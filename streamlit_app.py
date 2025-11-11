import pandas as pd
import streamlit as st

from loader import load_project, normalize_df
from elements import ElementFactory, Pricelist
from summarizer import Summarizer

# Show app title and description.
st.set_page_config(page_title="Passive Tools", page_icon="🛠️", layout="wide")
st.title("🛠️ Passive Tools")

# --- Step 1: File Upload ---
uploaded_file = st.file_uploader(
    "Choose a file",
    type=['csv', 'xlsx'],
)

if uploaded_file is not None:
    blueprints_df, manual_df, header_df = load_project(uploaded_file)
    df = normalize_df(blueprints_df, manual_df)
    
    header = header_df.iloc[0].to_dict()
    header["zakázka:"] = st.text_input("Zakázka", header.get("zakázka:"))
    header["stavba:"] = st.text_input("Stavba", header.get("stavba:"))
    header["č. zakázky:"] = st.text_input("Číslo zakázky", header.get("č. zakázky:"))
    
    st.success("Načteno")

    df = normalize_df(blueprints_df, manual_df)

    pricelist = Pricelist()
    elements = df.apply(lambda row: ElementFactory.create_element(row, pricelist), axis=1).tolist()

    elements_df = pd.DataFrame([e.to_dict() for e in elements])

    insulation_df = elements_df.groupby('insulation_mm')['insulation_area_m2'].sum().rename("quantity").to_frame().reset_index(names="spec")
    insulation_df.drop(labels=0, inplace=True)
    insulation_df['spec'] = insulation_df['spec'].apply(lambda x: f"tl={int(x)}")
    insulation_df[["system", "name", "unit", "position", "issues"]] = ["doplňkový a izolační materiál", "Izolace", "m2", "i", ""]

    elements_df = pd.concat([elements_df, insulation_df])

    s = Summarizer(header)
    s.write_inputs(blueprints_df, "Data z výkresu")
    s.write_inputs(manual_df, "Data doplněná")
    s.write_inputs(elements_df, "Data")
    s.write_inventory(elements_df)
    s.write_shopping_list(elements_df)

    # Create the download button
    st.download_button(
        label="💾 Download Excel File",
        data=s.close(),
        file_name=f"vypis_{uploaded_file.name.removesuffix('.csv').removesuffix('.xlsx')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    # Show instructions if no file has been uploaded yet
    st.info("Awaiting file upload to begin.")