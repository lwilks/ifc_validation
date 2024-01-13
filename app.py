import streamlit as st
#from IFCtoExcel import ifc_to_excel  # assuming you have a function named ifc_to_excel in IFCtoExcel.py
from IFCtoCSV import extract_csv  # assuming this is the function you want to use from IFCtoCSV.py

st.title('IFC Converter')

uploaded_file = st.sidebar.file_uploader("Choose an IFC file", type="ifc")
output_format = st.sidebar.selectbox('Choose output format', ('CSV'))

if uploaded_file is not None and output_format:
    if output_format == 'CSV':
        data = extract_csv(uploaded_file)  # replace 'YourClassType' with the class type you want
        st.write(data)
    #elif output_format == 'Excel':
        #excel_file = ifc_to_excel(uploaded_file)  # assuming ifc_to_excel returns an excel file
        #st.download_button(label="Download Excel file", data=excel_file, file_name='output.xlsx')