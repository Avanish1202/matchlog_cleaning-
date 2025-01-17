import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Custom CSS for orange background
st.markdown(
    """
    <style>
    body {
        background-color: orange;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Function to process Ageing data
def process_ageing(df):
    ageing_df = df[df['OUT DATE'].isnull()]
    columns_to_keep = ['SR. NO', 'CONTAINER NO', 'SIZE', 'SHIPPING LINE', 'IN DATE', 'GROSS WT', 'TARE WT']
    ageing_df = ageing_df[columns_to_keep]

    # Format 'IN DATE' to remove the time portion
    ageing_df['IN DATE'] = pd.to_datetime(ageing_df['IN DATE']).dt.strftime('%d-%m-%Y')

    # Calculate Ageing column
    try:
        ageing_df['Ageing'] = (pd.Timestamp.now() - pd.to_datetime(ageing_df['IN DATE'], format='%d-%m-%Y')).dt.days
    except Exception as e:
        st.error(f"Error calculating Ageing: {e}")
        ageing_df['Ageing'] = np.nan

    # Define bins for Ageing ranges
    bins = [-1, 3, 7, 15, 30, 60, 90, 180, float('inf')]
    labels = ['0-3', '04-07', '08-15', '16-30', '31-60', '61-90', '91-180', 'More than 180']
    ageing_df['Range'] = pd.cut(ageing_df['Ageing'], bins=bins, labels=labels, right=True)

    return ageing_df

# Function to process TAT data
def process_tat(df):
    tat_df = df.dropna(subset=['OUT DATE'])
    columns_to_keep = ['CONTAINER NO', 'SIZE', 'SHIPPING LINE', 'IN DATE', 'TRANSPORTERS', 'OUT DATE']
    tat_df = tat_df[columns_to_keep]

    # Format 'IN DATE' and 'OUT DATE' to remove the time portion
    tat_df['IN DATE'] = pd.to_datetime(tat_df['IN DATE']).dt.strftime('%d-%m-%Y')
    tat_df['OUT DATE'] = pd.to_datetime(tat_df['OUT DATE']).dt.strftime('%d-%m-%Y')

    # Calculate TAT column
    try:
        tat_df['TAT'] = (
            pd.to_datetime(tat_df['OUT DATE'], format='%d-%m-%Y') -
            pd.to_datetime(tat_df['IN DATE'], format='%d-%m-%Y')
        ).dt.days
    except Exception as e:
        st.error(f"Error calculating TAT: {e}")
        tat_df['TAT'] = np.nan

    # Define bins for TAT ranges
    bins = [-1, 0, 3, 7, 15, 30, 60, 90, 180, float('inf')]
    labels = ['0', '0-3', '04-07', '08-15', '16-30', '31-60', '61-90', '91-180', 'More than 180']
    tat_df['Range'] = pd.cut(tat_df['TAT'], bins=bins, labels=labels, right=True)

    return tat_df

# Function to create a downloadable Excel file
def generate_excel_file(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)  # Reset the pointer to the beginning of the buffer
    return output

# Streamlit UI
st.title("Ageing and TAT Sheet Generator")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel file
    df = pd.read_excel(uploaded_file)
    
    # Checkbox to select which sheet to process
    ageing_selected = st.checkbox("Generate Ageing Sheet")
    tat_selected = st.checkbox("Generate TAT Sheet")
    
    if ageing_selected:
        st.subheader("Ageing Sheet")
        ageing_df = process_ageing(df)
        st.dataframe(ageing_df)

        # Add download button for Ageing sheet
        ageing_excel = generate_excel_file(ageing_df)
        st.download_button(
            label="Download Ageing Sheet",
            data=ageing_excel,
            file_name="ageing.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    
    if tat_selected:
        st.subheader("TAT Sheet")
        tat_df = process_tat(df)
        st.dataframe(tat_df)

        # Add download button for TAT sheet
        tat_excel = generate_excel_file(tat_df)
        st.download_button(
            label="Download TAT Sheet",
            data=tat_excel,
            file_name="tat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if not ageing_selected and not tat_selected:
        st.info("Please select at least one sheet to generate.")
