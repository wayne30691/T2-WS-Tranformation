# T2‑WS Transformation

A streamlined Python + Streamlit solution for data transformation and visualization, designed to support structured reporting across trade channels.

## Project Structure

```
T2-WS-Transformation/
├── streamlit_app.py           # Main application entry point
├── requirements.txt           # Python dependencies
├── assets/                    # Screenshots and visuals
├── .github/                   # GitHub workflows
├── .devcontainer/             # Dev container setup (optional)
├── tests/                     # Unit tests (if any)
└── README.md
```

## Installation & Execution

All Python packages required for execution are listed in `requirements.txt`.  
The application runs with Streamlit at [T2 WS Transformation App](https://t2-ws-tranformation-php7ldvgfhtvkgpsaanptq.streamlit.app/).

---

## Features

- Clean UI built with **Streamlit** for quick insight delivery
- Modular Python architecture for **data preprocessing & transformation**
- Snapshot logic to prevent refresh failures caused by live-edited Excel files
- Designed to complement enterprise reporting tool in Pernod Ricard Taiwan

---

## Data Flow Overview

The application processes a combination of manually prepared inventory and sales reports, along with mapping tables retrieved from Salesforce via API. The transformation flow is structured as follows:

- **Input Layer**  
  Prepared Excel or CSV files containing sell-in / sell-out / stock-taking data  
  +  
  Mapping tables from Salesforce, connected via Power Query with API

- **Transformation Layer**  
  Python scripts perform data validation, cleaning, and merging with Salesforce mappings

- **Output Layer**  
  Processed results are exported as structured CSV files, which serve as clean inputs for enterprise reporting tool in Pernod Ricard Taiwan


