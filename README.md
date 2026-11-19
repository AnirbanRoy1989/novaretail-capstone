NovaRetail Q2 2026 Executive QBR

This project programmatically generates a fixed five-slide NovaRetail Quarterly Business Review presentation from the supplied Excel dataset and PowerPoint template.

The solution combines deterministic data processing with a Google Agent Development Kit (ADK) LlmAgent. Pandas calculates the KPIs and chart data, the ADK agent analyzes the results and produces structured executive insights, and python-pptx assembles the final presentation using the template’s named shapes.

Key capabilities
Loads and analyzes NovaRetail’s regional performance, customer experience, and operational risk data
Uses a Gemini-backed Google ADK agent to generate insight-led narrative
Verifies material claims through custom data-analysis tools
Validates the agent’s output with a Pydantic schema
Calculates KPI values deterministically with pandas
Creates editable, native PowerPoint charts
Finds and updates template elements by shape name
Preserves the supplied five-slide structure and layout
Produces a complete, presentation-ready PowerPoint file
Project files
generate_novaretail_deck_adk.py – main deck-generation script
novaretail_agent.py – Google ADK agent, tools, and output schema
novaretail_template.pptx – supplied fixed-layout PowerPoint template
novaretail_dataset.xlsx – supplied synthetic business dataset
requirements.txt – Python dependencies
NovaRetail_Q2_2026_QBR.pptx – generated executive presentation
Google ADK agent

The project defines a Google ADK LlmAgent named novaretail_qbr_agent.

The agent uses Gemini to generate:

Executive summary
Performance headline and insight bullets
Customer-experience headline and insight bullets
Operational-risk headline and insight bullets
Recommended actions
Closing leadership statement

The agent is instructed to use concise, insight-led executive language and avoid inventing values. Before writing the narrative, it can call the following custom tools:

get_metric

Retrieves an exact metric for a specified sheet, region, and quarter.

rank_regions

Ranks all regions by a selected metric for a given quarter.

compare_quarters

Calculates the absolute and percentage change in a regional metric between two quarters.

The agent returns a structured DeckNarrative object validated with Pydantic. This ensures that all required headlines, summaries, bullets, recommendations, and closing content are present before the PowerPoint is generated.

Architecture

The generation workflow is:

Load the three Excel worksheets with pandas.
Calculate the KPIs and chart series deterministically.
Make the datasets available to the ADK agent’s analysis tools.
Run the Gemini-backed LlmAgent.
Validate the agent response against the DeckNarrative Pydantic schema.
Open the fixed PowerPoint template.
Find each target using its assigned PowerPoint shape name.
Populate narrative text, KPI values, native charts, and icons.
Remove or replace instructional placeholder content.
Save the completed presentation as NovaRetail_Q2_2026_QBR.pptx.

The LLM is responsible for interpretation and executive wording. Numeric calculations, chart values, template positioning, and file generation remain deterministic.

Requirements
Python 3.10 or later
Google Gemini API key or a configured Google Cloud Vertex AI environment
PowerPoint template and Excel dataset in the project directory
