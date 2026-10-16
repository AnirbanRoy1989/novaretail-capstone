# NovaRetail Q2 2026 Executive QBR

This project programmatically updates the fixed five-slide NovaRetail template using the supplied Excel dataset. It uses each named template shape to place insight-led copy, KPI values, native PowerPoint charts and theme icons.

## Files
- `generate_novaretail_deck.py` — generation script
- `novaretail_template.pptx` — supplied fixed-layout template
- `novaretail_dataset.xlsx` — supplied synthetic source data
- `NovaRetail_Q2_2026_QBR.pptx` — generated PowerPoint

## Run
```bash
python -m pip install python-pptx pandas openpyxl
python generate_novaretail_deck.py
```

The output is written as `NovaRetail_Q2_2026_QBR.pptx` in the project folder.

## Design choices
- Slide 1: Q2 group revenue, sequential revenue growth and average regional NPS
- Slide 2: four-quarter regional revenue comparison
- Slide 3: four-quarter NPS trend
- Slide 4: Q2 stockout rate, supported by fraud, returns and capacity insights
- Slide 5: actions tied directly to APAC capacity pressure and Latin America deterioration

The brief referenced an icon repository that was not included in the supplied attachments. The script therefore uses clean native PowerPoint icon badges so the project remains self-contained and regenerates without external assets.
