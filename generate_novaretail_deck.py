from pathlib import Path
from datetime import date
import math
import pandas as pd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / 'novaretail_template.pptx'
DATA = ROOT / 'novaretail_dataset.xlsx'
OUTPUT = ROOT / 'NovaRetail_Q2_2026_QBR.pptx'

NAVY = RGBColor(30, 39, 97)
BLUE = RGBColor(48, 89, 167)
PERIWINKLE = RGBColor(96, 120, 190)
TEAL = RGBColor(0, 150, 136)
ORANGE = RGBColor(232, 126, 4)
RED = RGBColor(190, 52, 52)
GREY = RGBColor(91, 107, 140)
WHITE = RGBColor(255, 255, 255)
PALETTE = [BLUE, TEAL, ORANGE, RED]


def shape(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    raise KeyError(f'Missing named shape: {name}')


def remove_shape(slide, name):
    s = shape(slide, name)
    s._element.getparent().remove(s._element)


def set_text(s, text, size=None, colour=None, bold=None, italic=None, bullets=False):
    tf = s.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0
        if bullets:
            p.text = '• ' + p.text
        p.space_after = Pt(7 if bullets else 0)
        for r in p.runs:
            r.font.name = 'Aptos'
            if size: r.font.size = Pt(size)
            if colour: r.font.color.rgb = colour
            if bold is not None: r.font.bold = bold
            if italic is not None: r.font.italic = italic
    return s


def add_native_chart(slide, box, chart_type, categories, series, title=None, percent_axis=False):
    x, y, w, h = box.left, box.top, box.width, box.height
    remove_shape(slide, box.name)
    data = ChartData()
    data.categories = categories
    for name, values in series:
        data.add_series(name, values)
    chart = slide.shapes.add_chart(chart_type, x, y, w, h, data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.category_axis.tick_labels.font.size = Pt(10)
    chart.legend.font.size = Pt(10)
    chart.chart_style = 10
    if percent_axis:
        chart.value_axis.tick_labels.number_format = '0.0%'
        chart.value_axis.tick_labels.number_format_is_linked = False
    for idx, ser in enumerate(chart.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = PALETTE[idx % len(PALETTE)]
        ser.format.line.color.rgb = PALETTE[idx % len(PALETTE)]
    return chart


def add_icon(slide, placeholder_name, label, symbol, colour):
    ph = shape(slide, placeholder_name)
    x, y, w, h = ph.left, ph.top, ph.width, ph.height
    remove_shape(slide, placeholder_name)
    icon = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, min(w, h), min(w, h))
    icon.fill.solid(); icon.fill.fore_color.rgb = colour
    icon.line.color.rgb = colour
    tf = icon.text_frame; tf.clear(); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = symbol; r.font.name = 'Arial'; r.font.size = Pt(25); r.font.bold = True; r.font.color.rgb = WHITE
    icon.name = label


def main():
    perf = pd.read_excel(DATA, sheet_name='Regional_Performance')
    cx = pd.read_excel(DATA, sheet_name='Customer_Experience')
    risk = pd.read_excel(DATA, sheet_name='Operational_Risk')
    quarters = list(perf['Quarter'].drop_duplicates())
    regions = list(perf['Region'].drop_duplicates())
    latest, prior = quarters[-1], quarters[-2]

    latest_perf = perf[perf.Quarter == latest].set_index('Region')
    prior_perf = perf[perf.Quarter == prior].set_index('Region')
    latest_cx = cx[cx.Quarter == latest].set_index('Region')
    latest_risk = risk[risk.Quarter == latest].set_index('Region')

    revenue = latest_perf.Revenue_USD_M.sum()
    prior_revenue = prior_perf.Revenue_USD_M.sum()
    qoq = revenue / prior_revenue - 1
    avg_nps = latest_cx.NPS.mean()
    growth = (latest_perf.Revenue_USD_M / prior_perf.Revenue_USD_M - 1).sort_values(ascending=False)

    prs = Presentation(TEMPLATE)

    # Slide 1
    s = prs.slides[0]
    set_text(shape(s, 'subtitle_1'), 'Q2 2026 • Executive Review', 13, RGBColor(202,220,252), italic=False)
    set_text(shape(s, 'kpi1_value'), f'${revenue/1000:.2f}B', 28, WHITE, bold=True)
    set_text(shape(s, 'kpi1_label'), 'Q2 revenue', 12, RGBColor(202,220,252))
    set_text(shape(s, 'kpi2_value'), f'{qoq:+.1%}', 28, WHITE, bold=True)
    set_text(shape(s, 'kpi2_label'), 'Revenue vs Q1 2026', 12, RGBColor(202,220,252))
    set_text(shape(s, 'kpi3_value'), f'{avg_nps:.1f}', 28, WHITE, bold=True)
    set_text(shape(s, 'kpi3_label'), 'Average regional NPS', 12, RGBColor(202,220,252))
    summary = (f'NovaRetail closed Q2 at ${revenue/1000:.2f}B revenue, up {qoq:.1%} sequentially. '
               f'APAC led momentum with {growth.index[0]} revenue growth of {growth.iloc[0]:.1%} and the highest NPS at {latest_cx.loc["APAC","NPS"]}. '
               f'North America remained the largest market at ${latest_perf.loc["North America","Revenue_USD_M"]:.0f}M. '
               f'Latin America is the priority risk: revenue declined {abs(growth.loc["Latin America"]):.1%}, while returns, stockouts and fraud all worsened.')
    set_text(shape(s, 'summary_text'), summary, 17, WHITE)
    remove_shape(s, 'icon_1_label'); add_icon(s, 'icon_1', 'icon_overall', '↗', TEAL)

    # Slide 2
    s = prs.slides[1]
    set_text(shape(s, 'title_2'), 'APAC accelerates while Latin America contracts', 25, NAVY, bold=True, italic=False)
    b = shape(s, 'chart_performance')
    add_native_chart(s, b, XL_CHART_TYPE.COLUMN_CLUSTERED, regions,
                     [(q, [perf[(perf.Quarter==q)&(perf.Region==r)].Revenue_USD_M.iloc[0] for r in regions]) for q in quarters],
                     'Revenue by region (USD M)')
    remove_shape(s, 'chart_performance_label')
    bullets = [
        f'APAC delivered the strongest Q2 growth at {growth.loc["APAC"]:.1%} QoQ, reaching ${latest_perf.loc["APAC","Revenue_USD_M"]:.0f}M.',
        f'North America generated ${latest_perf.loc["North America","Revenue_USD_M"]:.0f}M and remains 36% of group revenue.',
        f'Europe recovered {growth.loc["Europe"]:.1%} QoQ to ${latest_perf.loc["Europe","Revenue_USD_M"]:.0f}M.',
        f'Latin America fell {abs(growth.loc["Latin America"]):.1%} QoQ; returns rose to {latest_perf.loc["Latin America","Returns_Rate_pct"]:.1f}% even as marketing spend decreased.'
    ]
    set_text(shape(s, 'insights_performance'), bullets, 15, NAVY, bullets=True)
    remove_shape(s, 'icon_2_label'); add_icon(s, 'icon_2', 'icon_growth', '↗', TEAL)

    # Slide 3
    s = prs.slides[2]
    set_text(shape(s, 'title_3'), 'Customer experience improves everywhere except Latin America', 24, NAVY, bold=True, italic=False)
    b = shape(s, 'chart_customer')
    add_native_chart(s, b, XL_CHART_TYPE.LINE_MARKERS, quarters,
                     [(r, list(cx[cx.Region==r].NPS)) for r in regions], 'NPS trend by region')
    remove_shape(s, 'chart_customer_label')
    bullets = [
        f'APAC leads at NPS {latest_cx.loc["APAC","NPS"]} and {latest_cx.loc["APAC","CSAT_pct"]:.0f}% CSAT, with the fastest delivery at {latest_cx.loc["APAC","Avg_Delivery_Days"]:.1f} days.',
        f'Europe improved to NPS {latest_cx.loc["Europe","NPS"]}, up 5 points since Q3 2025.',
        f'North America remains stable at NPS {latest_cx.loc["North America","NPS"]} and {latest_cx.loc["North America","CSAT_pct"]:.0f}% CSAT.',
        f'Latin America dropped to NPS {latest_cx.loc["Latin America","NPS"]}; delivery lengthened to {latest_cx.loc["Latin America","Avg_Delivery_Days"]:.1f} days and tickets reached {latest_cx.loc["Latin America","Support_Tickets_K"]:.0f}K.'
    ]
    set_text(shape(s, 'insights_customer'), bullets, 15, NAVY, bullets=True)
    remove_shape(s, 'icon_3_label'); add_icon(s, 'icon_3', 'icon_customer', '★', BLUE)

    # Slide 4
    s = prs.slides[3]
    set_text(shape(s, 'title_4'), 'Operational strain is concentrated in Latin America and APAC', 24, NAVY, bold=True, italic=False)
    b = shape(s, 'chart_risk')
    add_native_chart(s, b, XL_CHART_TYPE.BAR_CLUSTERED, regions,
                     [('Stockout rate (%)', [latest_risk.loc[r,'Stockout_Rate_pct'] for r in regions])],
                     'Q2 2026 stockout rate (%)')
    remove_shape(s, 'chart_risk_label')
    bullets = [
        f'Latin America has the highest stockout rate at {latest_risk.loc["Latin America","Stockout_Rate_pct"]:.1f}% and slowest returns at {latest_risk.loc["Latin America","Return_Processing_Days"]:.0f} days.',
        f'Fraud incidents in Latin America rose to {latest_risk.loc["Latin America","Fraud_Incidents"]:.0f}, versus 18 in Q3 2025.',
        f'APAC capacity reached {latest_risk.loc["APAC","Warehouse_Capacity_Util_pct"]:.0f}% and stockouts climbed to {latest_risk.loc["APAC","Stockout_Rate_pct"]:.1f}%, signalling a growth constraint.',
        f'North America and Europe are comparatively controlled, with Q2 stockouts at {latest_risk.loc["North America","Stockout_Rate_pct"]:.1f}% and {latest_risk.loc["Europe","Stockout_Rate_pct"]:.1f}%.'
    ]
    set_text(shape(s, 'insights_risk'), bullets, 14.5, NAVY, bullets=True)
    remove_shape(s, 'icon_4_label'); add_icon(s, 'icon_4', 'icon_risk', '!', RED)

    # Slide 5
    s = prs.slides[4]
    actions = [
        'EXPAND APAC CAPACITY: create a Q3 capacity plan to bring warehouse utilisation below 90% and arrest the rise in stockouts.',
        'STABILISE LATIN AMERICA FULFILMENT: launch a 30-day root-cause sprint across inventory allocation, carrier performance and return processing.',
        'TIGHTEN FRAUD CONTROLS: review high-risk transactions and payment rules in Latin America, where incidents reached 30 in Q2.',
        'REALLOCATE GROWTH INVESTMENT: protect APAC momentum while tying Latin America marketing spend to service-level recovery gates.'
    ]
    set_text(shape(s, 'recommendations_text'), actions, 18, WHITE, bullets=True)
    set_text(shape(s, 'closing_statement'), 'Leadership decision: fund capacity where growth is proven and make Latin America recovery the next-quarter operating priority.', 15, RGBColor(202,220,252), bold=True, italic=False)
    remove_shape(s, 'icon_5_label'); add_icon(s, 'icon_5', 'icon_action', '✓', ORANGE)

    prs.save(OUTPUT)
    print(OUTPUT)

if __name__ == '__main__':
    main()
