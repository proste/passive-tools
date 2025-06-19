import pandas as pd
import streamlit as st

from sheets import read_zwcad_dump, to_excel
from transformations import insulation_transformation

# Show app title and description.
st.set_page_config(page_title="Passive Tools", page_icon="🛠️", layout="wide")
st.title("🛠️ Passive Tools")
st.write(
    """
    Všechno nejlepší k narozeninám 🎉
    """
)

# --- Session State Initialization ---
# We use session state to hold the data across reruns
if 'edited_df' not in st.session_state:
    st.session_state.edited_df = pd.DataFrame()
if 'original_df' not in st.session_state:
    st.session_state.original_df = pd.DataFrame()
if 'file_name' not in st.session_state:
    st.session_state.file_name = ""


# --- Step 1: File Upload ---
st.header("1. Upload Your CSV File")
uploaded_file = st.file_uploader(
    "Choose a CSV file",
    type=['csv'],
)

# --- Main Logic ---
# This block runs when a file is uploaded
if uploaded_file is not None:
    # Store the filename for the download button
    st.session_state.file_name = uploaded_file.name

    try:
        # Load the uploaded file into a pandas DataFrame
        df = read_zwcad_dump(uploaded_file)
        st.session_state.original_df = df

        # --- Step 2: Apply Transformation ---
        st.header("2. Transform Data")
        with st.spinner('Applying transformation...'):
            # transformed_df = df
            transformed_df = insulation_transformation(df.copy()) # Use a copy to be safe

        st.success("Transformation applied successfully!")

        # --- Step 3: Display and Edit Data ---
        st.header("3. Review and Edit Transformed Data")
        st.write("You can edit the values directly in the table below. Changes are saved automatically.")

        # The data_editor widget allows direct editing of the DataFrame
        # The key is important to keep the state of the editor consistent
        edited_df = st.data_editor(transformed_df, num_rows="dynamic", key="data_editor")
        st.session_state.edited_df = edited_df

        summary_df = st.dataframe(transformed_df.groupby("insulation_thickness").insulation_area.sum().drop(0.0).round(2))

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")


# --- Step 4: Download Corrected Data ---
# This section is visible only after data has been processed
if not st.session_state.edited_df.empty:
    st.header("4. Download Corrected File")
    st.write("Click the button below to download the edited data as an Excel file.")

    # Convert the edited DataFrame to Excel format
    excel_data = to_excel(st.session_state.edited_df)

    # Create the download button
    st.download_button(
        label="💾 Download Excel File",
        data=excel_data,
        file_name=f"transformed_{st.session_state.file_name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    # Show instructions if no file has been uploaded yet
    st.info("Awaiting file upload to begin.")