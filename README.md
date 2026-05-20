# Excel Sheet Comparator & Merger (Streamlit)

A lightweight Streamlit app that compares columns across multiple Excel sheets/files and exports the results.

## Features
- Upload **multiple Excel workbooks (.xlsx)**
- Select **multiple sheets** to compare
- Choose **specific columns** from each sheet
- Get:
  - **Union of selected columns**
  - **Common columns**
  - **Columns only in the first selected sheet**
  - **Column presence matrix**
  - **Merged output** (all rows appended after aligning columns)
- Download the output as **Excel** or **CSV**

## Typical Example
If:
- Sheet 1 has columns: `A, B, C`
- Sheet 2 has columns: `B, C, D, E`

The app can show:
- **Union:** `A, B, C, D, E`
- **Common:** `B, C`
- **Only in first sheet:** `A`
- **Merged output:** rows from both sheets aligned to the union of selected columns.

> Note: In your example you mentioned an output like `A, B, C, D`. If you want to exclude `E`, you can simply leave `E` unselected in the column picker.

## Project structure
```text
.
├── app.py
├── requirements.txt
└── README.md
```

## Run locally
1. Create and activate a virtual environment (recommended)
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Start the app:
   ```bash
   streamlit run app.py
   ```

## Deploy on Streamlit Community Cloud
1. Push these files to a GitHub repository.
2. Go to **https://share.streamlit.io/**
3. Click **New app**
4. Select your GitHub repo, branch, and `app.py`
5. Deploy

## Suggested GitHub repo name
`excel-sheet-comparator-streamlit`

## Future enhancements (optional)
- Fuzzy column matching (e.g., `Client Name` vs `Client_Name`)
- Compare data values, not just headers
- Support `.xls` files
- Rule-based output (e.g., include only common + unique from first file)
- Save column selection templates
