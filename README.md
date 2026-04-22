# 📊 Excel File Comparator

A Streamlit web app that compares two Excel files cell-by-cell and highlights matches vs mismatches in a downloadable report.

## Live Demo

> Deploy link will appear here after Streamlit Cloud deployment.

## Features

- Upload any two `.xlsx` files and compare them instantly
- **Green** cells = values match · **Yellow** cells = values differ
- Optional **key column** for row matching (e.g. ID, OrderNo) instead of positional matching
- **Case-sensitive / insensitive** comparison toggle
- **Mismatch preview table** in the browser before downloading
- Summary stats — total cells, matches, mismatches, and match rate %
- Download a fully colour-coded Excel report

## How to Use

1. Upload the **Correct File** (the reference/expected file)
2. Upload the **Incorrect File** (the file to check)
3. *(Optional)* Set a key column and comparison options
4. Click **Compare & Generate Report**
5. Review the mismatch preview and download the report

## Run Locally

```bash
git clone https://github.com/sudhir-yadav28/Excel_Compare.git
cd Excel_Compare
pip install -r requirements.txt
streamlit run app.py
```

## Tech Stack

- [Streamlit](https://streamlit.io/) — UI framework
- [Pandas](https://pandas.pydata.org/) — data comparison
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel read/write with cell styling
