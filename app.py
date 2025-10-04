import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="SAP vs Open Orders — 4-Field Row-by-Row Checker", layout="wide")

# ---------- Helpers ----------
def clean_model(x: str) -> str:
    if pd.isna(x):
        return ""
    s = str(x)
    return s.split(">>", 1)[0].strip() if ">>" in s else s.strip()

def norm_po(x: str) -> str:
    return str(x).strip().upper().replace(" ", "")

def norm_line(x: str) -> str:
    s = re.sub(r"[^0-9]", "", str(x))
    return s.lstrip("0") or "0"

def norm_loc_sap(x) -> str:
    code = str(x).strip()
    return "EDMONTON" if code == "2913" else code.upper()

def norm_loc_oo(x) -> str:
    return str(x).strip().upper()

def extract_sap_four(df: pd.DataFrame) -> pd.DataFrame:
    sap = df.iloc[:, [2, 3, 6, 7]].copy()
    sap.columns = ["po", "location", "po_item", "model"]
    sap["po"] = sap["po"].apply(norm_po)
    sap["po_item"] = sap["po_item"].apply(norm_line)
    sap["model"] = sap["model"].apply(clean_model)
    sap["location"] = sap["location"].apply(norm_loc_sap)
    return sap

def extract_oo_four(df: pd.DataFrame) -> pd.DataFrame:
    oo = df.iloc[:, [13, 14, 16, 9]].copy()
    oo.columns = ["po", "po_item", "model", "location"]
    oo["po"] = oo["po"].apply(norm_po)
    oo["po_item"] = oo["po_item"].apply(norm_line)
    oo["model"] = oo["model"].apply(clean_model)
    oo["location"] = oo["location"].apply(norm_loc_oo)
    return oo

def compare_row_by_row(sap_four: pd.DataFrame, oo_four: pd.DataFrame):
    n = min(len(sap_four), len(oo_four))
    sap_rows = sap_four.head(n).reset_index(drop=True)
    oo_rows = oo_four.head(n).reset_index(drop=True)

    rows_checked = n
    equivalent_rows = 0
    discrepancy_rows = []

    combined_rows = []
    for i in range(n):
        issues = []
        sap_row = sap_rows.loc[i]
        oo_row = oo_rows.loc[i]

        if sap_row["po"] != oo_row["po"]:
            issues.append(f"PO #: SAP={sap_row['po']} vs OO={oo_row['po']}")
        if sap_row["po_item"] != oo_row["po_item"]:
            issues.append(f"PO Item: SAP={sap_row['po_item']} vs OO={oo_row['po_item']}")
        if sap_row["model"] != oo_row["model"]:
            issues.append(f"Model: SAP='{sap_row['model']}' vs OO='{oo_row['model']}'")
        if sap_row["location"] != oo_row["location"]:
            issues.append(f"Location: SAP='{sap_row['location']}' vs OO='{oo_row['location']}'")

        combined_rows.append({
            "Excel Row": i + 2,
            "SAP — PO": sap_row["po"],
            "SAP — Item": sap_row["po_item"],
            "SAP — Model": sap_row["model"],
            "SAP — Location": sap_row["location"],
            "OO — PO": oo_row["po"],
            "OO — Item": oo_row["po_item"],
            "OO — Model": oo_row["model"],
            "OO — Location": oo_row["location"],
            "Equivalent?": "✅" if not issues else "❌"
        })

        if issues:
            discrepancy_rows.append({
                "Excel Row": i + 2,
                "What’s Different": " | ".join(issues)
            })
        else:
            equivalent_rows += 1

    combined_df = pd.DataFrame(combined_rows)
    discrepancy_df = pd.DataFrame(discrepancy_rows)
    return rows_checked, equivalent_rows, discrepancy_df, combined_df, len(sap_four), len(oo_four)

# ---------- UI ----------
st.title("SAP vs Open Orders — 4-Field Row-by-Row Comparison")
st.caption("Compares **PO #, PO Item, Model, Location** — strictly by Excel row order. Model is trimmed at `>>`. SAP location code `2913` is normalized to `EDMONTON`.")

left, right = st.columns(2)
with left:
    sap_file = st.file_uploader("Upload **SAP Excel** file", type=["xlsx", "xls"])
with right:
    oo_file = st.file_uploader("Upload **Open Orders Excel** file", type=["xlsx", "xls"])

run = st.button("Compare")

if run:
    if not sap_file or not oo_file:
        st.warning("Please upload both files before comparing.")
        st.stop()

    try:
        sap_df = pd.read_excel(sap_file)
        oo_df = pd.read_excel(oo_file)
    except Exception as e:
        st.error(f"Failed to read Excel files: {e}")
        st.stop()

    try:
        sap_four = extract_sap_four(sap_df)
        oo_four  = extract_oo_four(oo_df)
    except Exception as e:
        st.error(f"Column extraction failed. Error: {e}")
        st.stop()

    rows_checked, rows_equivalent, disc_df, combined_df, sap_len, oo_len = compare_row_by_row(sap_four, oo_four)

    st.subheader("Summary")
    a, b, c = st.columns(3)
    with a:
        st.metric(label="Rows Checked", value=rows_checked)
    with b:
        st.metric(label="Rows Equivalent", value=rows_equivalent)
        st.write("✅ Equivalent rows")
    with c:
        st.metric(label="Rows with Discrepancies", value=int(rows_checked - rows_equivalent))
        st.write("❌ Has at least one mismatch")

    if sap_len != oo_len:
        st.info(f"Files differ in length. SAP rows: **{sap_len}**, OO rows: **{oo_len}**. "
                f"Only the first **{rows_checked}** rows were compared.")

    st.divider()
    st.subheader("Row-by-Row View")
    st.dataframe(combined_df, use_container_width=True)

    st.subheader("Discrepancy Report")
    if disc_df.empty:
        st.success("No discrepancies found.")
    else:
        st.dataframe(disc_df, use_container_width=True)
        csv_bytes = disc_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download Discrepancies CSV",
            data=csv_bytes,
            file_name="discrepancies.csv",
            mime="text/csv"
        )
else:
    st.info("Upload both Excel files and click **Compare**.")
