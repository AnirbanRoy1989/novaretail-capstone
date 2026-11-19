# NovaRetail Q2 2026 Executive QBR

This project programmatically generates a fixed five-slide NovaRetail Quarterly Business Review presentation from the supplied Excel dataset and PowerPoint template.

The solution combines deterministic data processing with a Google Agent Development Kit (ADK) `LlmAgent`. Pandas calculates the KPIs and chart data, the ADK agent analyzes the results and produces structured executive insights, and `python-pptx` assembles the final presentation using the template's named shapes.

## Key capabilities

- Loads and analyzes NovaRetail's regional performance, customer experience, and operational risk data
- Uses a Gemini-backed Google ADK agent to generate insight-led narrative
- Verifies material claims through custom data-analysis tools
- Validates the agent's output with a Pydantic schema
- Calculates KPI values deterministically with pandas
- Creates editable, native PowerPoint charts
- Finds and updates template elements by shape name
- Preserves the supplied five-slide structure and layout
- Produces a complete, presentation-ready PowerPoint file

## Project files

- `generate_novaretail_deck_adk.py` - main deck-generation script
- `novaretail_agent.py` - Google ADK agent, tools, and output schema
- `requirements.txt` - Python dependencies
- `NovaRetail_Q2_2026_QBR.pptx` - generated executive presentation

## Google ADK agent

The project defines a Google ADK `LlmAgent` named `novaretail_qbr_agent`.

The agent uses Gemini to generate:

- Executive summary
- Performance headline and insight bullets
- Customer-experience headline and insight bullets
- Operational-risk headline and insight bullets
- Recommended actions
- Closing leadership statement

The agent is instructed to use concise, insight-led executive language and avoid inventing values. Before writing the narrative, it can call these custom tools:

### `get_metric`

Retrieves an exact metric for a specified sheet, region, and quarter.

### `rank_regions`

Ranks all regions by a selected metric for a given quarter.

### `compare_quarters`

Calculates the absolute and percentage change in a regional metric between two quarters.

The agent returns a structured `DeckNarrative` object validated with Pydantic. This ensures that all required headlines, summaries, bullets, recommendations, and closing content are present before the PowerPoint is generated.

## Architecture

The generation workflow is:

1. Load the three Excel worksheets with pandas.
2. Calculate the KPIs and chart series deterministically.
3. Make the datasets available to the ADK agent's analysis tools.
4. Run the Gemini-backed `LlmAgent`.
5. Validate the response against the `DeckNarrative` Pydantic schema.
6. Open the fixed PowerPoint template.
7. Find each target using its assigned PowerPoint shape name.
8. Populate narrative text, KPI values, native charts, and icons.
9. Remove or replace instructional placeholder content.
10. Save the completed presentation as `NovaRetail_Q2_2026_QBR.pptx`.

The LLM is responsible for interpretation and executive wording. Numeric calculations, chart values, template positioning, and file generation remain deterministic.

## Requirements

- Python 3.10 or later
- Google Gemini API key or a configured Google Cloud Vertex AI environment
- PowerPoint template and Excel dataset in the project directory

Install the dependencies:

```bash
python -m pip install -r requirements.txt
```

Alternatively:

```bash
python -m pip install google-adk pydantic python-pptx pandas openpyxl
```

## Authentication

### Gemini Developer API

Set the `GOOGLE_API_KEY` environment variable.

macOS or Linux:

```bash
export GOOGLE_API_KEY="your-api-key"
```

Windows PowerShell:

```powershell
$env:GOOGLE_API_KEY="your-api-key"
```

### Vertex AI

To use Vertex AI instead of the Gemini Developer API, configure:

```bash
export GOOGLE_GENAI_USE_VERTEXAI="TRUE"
export GOOGLE_CLOUD_PROJECT="your-project-id"
export GOOGLE_CLOUD_LOCATION="your-location"
```

The script stops with a clear error if neither Gemini API nor Vertex AI authentication is configured. It does not silently replace the agent output with hard-coded narrative.

## Model configuration

The default model is:

```text
gemini-2.5-flash
```

To use another ADK-compatible model, set `NOVARETAIL_MODEL` before running the generator.

macOS or Linux:

```bash
export NOVARETAIL_MODEL="gemini-2.5-flash"
```

Windows PowerShell:

```powershell
$env:NOVARETAIL_MODEL="gemini-2.5-flash"
```

## Run the project

Ensure these files are in the same project directory:

```text
generate_novaretail_deck_adk.py
novaretail_agent.py
novaretail_template.pptx
novaretail_dataset.xlsx
requirements.txt
```

Generate the presentation:

```bash
python generate_novaretail_deck_adk.py
```

The completed presentation is written to:

```text
NovaRetail_Q2_2026_QBR.pptx
```

## Slide design choices

1. **Executive overview** - Presents Q2 group revenue, sequential revenue growth, average regional NPS, and an agent-generated executive summary.
2. **Regional performance** - Compares four quarters of revenue by region and highlights the strongest growth, largest market, recovery patterns, and underperforming regions.
3. **Customer experience** - Shows the four-quarter NPS trend and connects it with CSAT, delivery time, and support-ticket performance.
4. **Operational risk** - Compares Q2 stockout rates and supplements the chart with fraud, returns-processing, and warehouse-capacity insights.
5. **Recommended actions** - Converts the performance and risk findings into concise actions, particularly around APAC capacity pressure and Latin America's operational deterioration.

## Icons

The original brief referenced an icon repository, but that repository was not included in the supplied project attachments. The generator therefore uses clean, native PowerPoint icon badges.

This keeps the project:

- Self-contained
- Fully reproducible
- Independent of external image services
- Editable in PowerPoint

If the original icon repository becomes available, the native badges can be replaced by selecting the most appropriate icon for each slide and inserting it into the corresponding named icon placeholder.

## ADK Web testing

The agent exports a `root_agent`, allowing it to be tested through the ADK development interface.

Place the agent module in a valid ADK application package as `agent.py`, then run:

```bash
adk web
```

This allows the agent and its function-tool calls to be inspected interactively before generating the final presentation.

## Output

The final deliverable is:

```text
NovaRetail_Q2_2026_QBR.pptx
```

It contains the required five slides, native editable charts, computed KPIs, and evidence-grounded executive narrative generated through Google ADK and Gemini.
