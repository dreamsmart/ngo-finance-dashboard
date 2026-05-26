# NGO Finance Dashboard

Initial Streamlit project structure for an NGO finance dashboard.

## Features

- Upload CSV or Excel transaction files
- Clean and normalize transaction columns
- Categorize transactions with YAML keyword rules
- Store processed transactions in SQLite
- Visualize income, expenses, categories, projects, and monthly trends
- Export cleaned data to Excel

## Project Structure

```text
.
├── app.py
├── requirements.txt
├── README.md
├── config
│   ├── category_rules.yaml
│   └── project_rules.yaml
├── data
│   ├── input
│   └── output
├── database
└── src
    ├── categorizer.py
    ├── cleaner.py
    ├── dashboard.py
    ├── export.py
    └── loader.py
```

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Expected Data

The app works best with transaction files containing columns similar to:

- `date`
- `description`
- `amount`
- `account`
- `donor`

Column names are normalized during cleaning, so common variations are accepted.
