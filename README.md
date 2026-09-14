# Budget Intelligence

Streamlit MVP for a new budget generation workflow.

The app starts from two evidence sets:

- Approved annual budget from last year
- Actuals YTD by department, cost center, account and category

It normalizes the inputs, builds a run-rate baseline, flags lines that need challenge, generates department-specific questions, captures owner answers and exports a proposed budget.

## Run

Use the Streamlit environment from the previous budget generator, or any Python environment with `streamlit`, `pandas`, `altair` and optionally `openpyxl`.

```bash
streamlit run app.py --server.port 8502
```

CSV uploads are supported by default. XLSX uploads require `openpyxl`.
