# Graduate Portfolio Intelligence — V1 Prototype

This is a small Streamlit prototype built from `Persistence_Model_Graduate_Programs__9_23_2026__3.xlsx`.

## What it currently does

- Executive portfolio KPIs
- Program-level persistence table
- Program explorer
- Method A / Method B persistence cross-check visibility
- Revenue/LTR snapshot
- Trial CAC snapshot where available

## Important model guardrails

- The Excel workbook remains the reference model.
- Trial/approximate sheets are not silently promoted to decision-grade logic.
- Graduation is not treated as observed completion because the workbook states the source is an enrolled-student report and cannot distinguish graduation from withdrawal.
- CAC is shown as directional because the workbook itself flags the CAC/LTR trial as needing validation.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit provides.

## Next build step

Validate the web outputs against representative Excel programs before adding scenario modeling, automated data refresh, or production deployment.
