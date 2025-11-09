import datetime as dt

import pandas as pd
import streamlit as st

from summarizer import insulate_df, normalize_df, summarize_df

# Show app title and description.
st.set_page_config(page_title="Passive Tools", page_icon="🛠️", layout="wide")
st.title("🛠️ Passive Tools")

# --- Step 1: File Upload ---
order = st.text_input("Zakázka", "")
parcel = st.text_input("Stavba", "")
order_id = st.text_input("Číslo zakázky", "")
uploaded_file = st.file_uploader(
    "Choose a file",
    type=['csv', 'xlsx'],
)

# --- Main Logic ---
# This block runs when a file is uploaded
if uploaded_file is not None:
    if uploaded_file.name.endswith("csv"):
        df = blueprints_df = pd.read_csv(uploaded_file, delimiter=";", encoding="cp1250", decimal=",")
        manual_df = pd.DataFrame([], columns=["Systém", "Číslo", "Název", "Typ", "Součet", "--", "PN"])
    else:  # uploaded_file.name.endswith("xlsx")
        blueprints_df =  pd.read_excel(uploaded_file, sheet_name="Data z výkresu")
        manual_df =  pd.read_excel(uploaded_file, sheet_name="Data doplněná")
        df = pd.concat([blueprints_df, manual_df]).reset_index()
    st.success("Načteno")

    with st.spinner('Izoluji...'):
        # transformed_df = df
        transformed_df = insulate_df(normalize_df(df.copy()))
    st.success("Zaizolováno")

    # Convert the edited DataFrame to Excel format
    with st.spinner('Shrnuji...'):
        header = {
            "zakázka:": order,
            "stavba:": parcel,
            "č. zakázky:": order_id,
            "vypracoval:": "M. Jindráková",
            "dne:": dt.date.today().strftime("%-d/%-m/%Y"),
        }
        summary_xlsx = summarize_df(blueprints_df, manual_df, transformed_df, header)
    st.success("Shrnuto")

    # Create the download button
    st.download_button(
        label="💾 Download Excel File",
        data=summary_xlsx,
        file_name=f"vypis_{uploaded_file.name.removesuffix('.csv').removesuffix('.xlsx')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    # Show instructions if no file has been uploaded yet
    st.info("Awaiting file upload to begin.")