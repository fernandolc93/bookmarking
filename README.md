# aCRF Automatic Bookmarking Tool

A Streamlit application that automatically generates a two-level PDF bookmark tree  
(Visit → Form) for annotated Case Report Forms (aCRFs) used in clinical trials.

## Features

- Upload an aCRF PDF and extract all form titles automatically
- Map forms to study visits via an interactive Form × Visit matrix
- Bulk-tick toolbar for fast matrix editing
- Excel import/export for the Form × Visit matrix
- Live coverage warnings for unmapped forms or empty visits
- Download the bookmarked PDF ready for regulatory submission

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Project Structure

```
├── app.py                  # Main Streamlit application
├── generate_matrix.py      # Standalone matrix template generator
├── requirements.txt
└── core/
    ├── extract.py          # PDF title and page extraction
    ├── plan.py             # Bookmark plan builder
    ├── pdf_writer.py       # pikepdf bookmark writer
    └── matrix_io.py        # Excel import/export for the matrix
```

## Deployment

Deployed on [Streamlit Community Cloud](https://share.streamlit.io).  
No PDF data is stored — all processing is in-memory per session.
