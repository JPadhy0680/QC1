
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Sheet Comparator", layout="wide")

st.title("Excel Sheet Comparator & Merger")
st.caption(
    "Upload one or more Excel files, select the sheets and columns you want to compare, "
    "and export the results."
)


def load_workbook(file_obj):
    """Return a dict of sheet_name -> DataFrame."""
    xls = pd.ExcelFile(file_obj, engine="openpyxl")
    sheets = {}
    for s in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=s, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        sheets[s] = df
    return sheets


def column_presence_map(selected_sources, selected_columns_map):
    all_cols = sorted(set().union(*[set(cols) for cols in selected_columns_map.values()])) if selected_columns_map else []
    rows = []
    for col in all_cols:
        row = {"Column": col}
        for source in selected_sources:
            source_name = source["label"]
            row[source_name] = "Yes" if col in selected_columns_map[source_name] else "No"
        rows.append(row)
    return pd.DataFrame(rows)


def get_safe_source_column_name(selected_sources, selected_columns_map):
    """Pick a metadata column name that won't clash with user data."""
    used_columns = set()
    for source in selected_sources:
        used_columns.update(selected_columns_map.get(source["label"], []))
        used_columns.update([str(c).strip() for c in source["df"].columns])

    candidates = [
        "Source",
        "Source_Sheet",
        "__Source__",
        "__Source_Sheet__",
        "_Merged_Source_",
    ]
    for candidate in candidates:
        if candidate not in used_columns:
            return candidate

    i = 1
    while True:
        candidate = f"_Merged_Source_{i}_"
        if candidate not in used_columns:
            return candidate
        i += 1


def build_merged_dataframe(selected_sources, selected_columns_map, source_column_name):
    """Concatenate selected sheets after aligning selected columns."""
    union_cols = list(dict.fromkeys(
        [col for source in selected_sources for col in selected_columns_map[source["label"]]]
    ))

    merged_parts = []
    for source in selected_sources:
        source_name = source["label"]
        df = source["df"].copy()
        picked = selected_columns_map[source_name]
        part = df[picked].copy()
        for c in union_cols:
            if c not in part.columns:
                part[c] = pd.NA
        part = part[union_cols]
        part.insert(0, source_column_name, source_name)
        merged_parts.append(part)

    if merged_parts:
        return pd.concat(merged_parts, ignore_index=True)
    return pd.DataFrame(columns=[source_column_name] + union_cols)


def as_excel_bytes(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs_dict.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output.getvalue()


with st.sidebar:
    st.header("Upload Excel files")
    uploaded_files = st.file_uploader(
        "Select one or more .xlsx files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="You can upload multiple Excel workbooks."
    )

if not uploaded_files:
    st.info("Upload Excel files from the left sidebar to begin.")
    st.stop()

workbooks = {}
read_errors = []
for file in uploaded_files:
    try:
        workbooks[file.name] = load_workbook(file)
    except Exception as e:
        read_errors.append(f"{file.name}: {e}")

if read_errors:
    st.error("Some files could not be read:")
    for err in read_errors:
        st.write(f"- {err}")

if not workbooks:
    st.stop()

all_sources = []
for file_name, sheets in workbooks.items():
    for sheet_name, df in sheets.items():
        label = f"{file_name} | {sheet_name}"
        all_sources.append({"file": file_name, "sheet": sheet_name, "label": label, "df": df})

st.subheader("1) Select sheets to compare")
source_labels = [s["label"] for s in all_sources]
default_labels = source_labels[:2] if len(source_labels) >= 2 else source_labels
chosen_labels = st.multiselect(
    "Choose at least two sheets",
    options=source_labels,
    default=default_labels
)

selected_sources = [s for s in all_sources if s["label"] in chosen_labels]

if len(selected_sources) < 2:
    st.warning("Please choose at least two sheets for comparison.")
    st.stop()

st.subheader("2) Choose columns from each selected sheet")
selected_columns_map = {}
cols_preview = st.columns(len(selected_sources))

for i, source in enumerate(selected_sources):
    with cols_preview[i]:
        st.markdown(f"**{source['label']}**")
        options = list(source["df"].columns)
        selected_cols = st.multiselect(
            f"Columns for {source['sheet']}",
            options=options,
            default=options,
            key=f"cols_{source['label']}"
        )
        if not selected_cols:
            st.error("Select at least one column.")
            st.stop()
        selected_columns_map[source["label"]] = selected_cols
        st.caption(f"Rows: {len(source['df'])} | Columns: {len(options)}")

source_column_sets = {name: set(cols) for name, cols in selected_columns_map.items()}
union_cols = sorted(set().union(*source_column_sets.values()))
common_cols = sorted(set.intersection(*source_column_sets.values()))

first_name = selected_sources[0]["label"]
first_set = source_column_sets[first_name]
others_union = set().union(*[s for n, s in source_column_sets.items() if n != first_name])
only_in_first = sorted(first_set - others_union)
only_in_others = sorted(others_union - first_set)

presence_df = column_presence_map(selected_sources, selected_columns_map)
source_column_name = get_safe_source_column_name(selected_sources, selected_columns_map)
merged_df = build_merged_dataframe(selected_sources, selected_columns_map, source_column_name)

st.subheader("3) Results")
metric_cols = st.columns(4)
metric_cols[0].metric("Sheets selected", len(selected_sources))
metric_cols[1].metric("Union columns", len(union_cols))
metric_cols[2].metric("Common columns", len(common_cols))
metric_cols[3].metric("Merged rows", len(merged_df))

result_tab1, result_tab2, result_tab3, result_tab4 = st.tabs([
    "Union / Common",
    f"Only in {selected_sources[0]['sheet']}",
    "Presence Matrix",
    "Merged Data",
])

with result_tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Union of Selected Columns")
        st.dataframe(pd.DataFrame({"Union Columns": union_cols}), use_container_width=True)
    with c2:
        st.markdown("### Common Columns")
        st.dataframe(pd.DataFrame({"Common Columns": common_cols}), use_container_width=True)

with result_tab2:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"### Columns only in {selected_sources[0]['label']}")
        st.dataframe(pd.DataFrame({"Only in First Selected Sheet": only_in_first}), use_container_width=True)
    with c2:
        st.markdown("### Columns available in other selected sheets but not in the first")
        st.dataframe(pd.DataFrame({"Only in Other Selected Sheets": only_in_others}), use_container_width=True)

with result_tab3:
    st.markdown("### Column Presence Matrix")
    st.dataframe(presence_df, use_container_width=True)

with result_tab4:
    st.markdown("### Normalized Merged Output")
    st.caption(
        f"A safe metadata column named '{source_column_name}' is added to indicate the origin sheet. "
        "This avoids clashes if your Excel file already contains a column named 'Source'."
    )
    st.dataframe(merged_df, use_container_width=True, height=450)

st.subheader("4) Download the output")
output_book = {
    "Summary_Union": pd.DataFrame({"Union Columns": union_cols}),
    "Summary_Common": pd.DataFrame({"Common Columns": common_cols}),
    "Presence_Matrix": presence_df,
    "Merged_Data": merged_df,
}
excel_data = as_excel_bytes(output_book)

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        label="Download Results as Excel",
        data=excel_data,
        file_name="excel_comparison_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    csv_data = merged_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Merged Data as CSV",
        data=csv_data,
        file_name="merged_data.csv",
        mime="text/csv",
        use_container_width=True,
    )

with st.expander("How this works"):
    st.markdown(
        """
        - Upload one or more `.xlsx` files.
        - Select the sheets you want to compare.
        - Select the columns from each sheet.
        - See the **union**, **common columns**, **presence matrix**, and **merged output**.
        - Download the final output in Excel or CSV format.

        **Example:**
        - Sheet 1: `A, B, C`
        - Sheet 2: `B, C, D, E`

        The tool can show:
        - **Union:** `A, B, C, D, E`
        - **Common:** `B, C`
        - **Only in first sheet:** `A`
        - **Merged output:** rows from both sheets aligned to the union of selected columns.
        """
    )
