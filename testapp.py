import streamlit as st
import pandas as pd

st.title("Baca Excel di Streamlit - pandas.read_excel")

uploaded = st.file_uploader("Upload file Excel (.xlsx/.xls)", type=["xlsx", "xls"])
if uploaded:
    # pandas bisa menerima file-like object langsung
    df = pd.read_excel(uploaded)            # untuk xlsx, engine=openpyxl akan dipakai otomatis
    st.dataframe(df)                        # preview
    st.write(df.head())
