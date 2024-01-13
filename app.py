import streamlit as st
from IFCtoCSV import extract_csv
from IFCtoExcel import extract_xlsx
import pandas as pd
import io

st.title('IFC Converter')

uploaded_file = st.sidebar.file_uploader("Choose an IFC file", type="ifc")
output_format = st.sidebar.selectbox('Choose output format', ('CSV','Excel'))

if uploaded_file is not None and output_format:
    file_name = uploaded_file.name.split('.')[0]  # use the original file name, but change the extension to .csv

    if output_format == 'CSV':
        data = extract_csv(uploaded_file)
        csv = data.to_csv(index=False).encode()
        st.download_button(label="Download CSV File", data=csv, file_name=file_name + '.csv', mime='text/csv')
    elif output_format == 'Excel':
        data = extract_xlsx(uploaded_file)
        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
            for object_class in data["Class"].unique():
                print(data)
                df_class = data[data["Class"] == object_class]
                df_class = df_class.dropna(axis=1, how='all')
                df_class.to_excel(writer, sheet_name=object_class, index=False)
        towrite.seek(0)  # go to the start of the BytesIO object
        excel = towrite.read()  # read the BytesIO object to get the Excel data
        st.download_button(label="Download Excel File", data=excel, file_name=file_name + '.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')