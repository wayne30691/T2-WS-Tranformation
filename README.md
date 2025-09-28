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

Install dependencies:

```bash
pip install -r requirements.txt
```

Launch the application:

```bash
streamlit run streamlit_app.py
```

The app is available at `[http://localhost:8501](https://t2-ws-tranformation-php7ldvgfhtvkgpsaanptq.streamlit.app/)`.

---

## Features

- Clean UI built with **Streamlit** for quick insight delivery
- Modular Python architecture for **data preprocessing & transformation**
- Snapshot logic to prevent refresh failures caused by live-edited Excel files
- Designed to complement enterprise reporting tool in Pernod Ricard Taiwan

---

## Data Flow Overview

- **Raw Data** (Excel, CSV, SharePoint)  
  ⮕ **Preprocessing Layer** (Python scripts / Power Query)  
  ⮕ **Snapshot Files** (staged for refresh)  
  ⮕ **Visualization** (Streamlit or Power BI dashboards)  
  ⮕ **Governance** (role-based access control aligned with HQ policy)
