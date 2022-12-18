import streamlit as st  # pip install streamlit
import pandas as pd  # pip install pandas
import plotly.express as px  # pip install plotly-express
import base64  # Standard Python Module
from io import StringIO, BytesIO  # Standard Python Module


def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Ladda ner Excel-fil</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_plot(df_bokning, column, xaxis_label, yaxis_label, ti):
    with column:
        fig = px.bar(
            df_bokning, 
            x = df_bokning.columns[0], 
            y = df_bokning.columns[-1], 
            template='seaborn',
            labels = {
                df_bokning.columns[0]: xaxis_label,
                df_bokning.columns[-1]: yaxis_label
            }, 
            title = ti)
        st.plotly_chart(fig, use_container_width = True)

def display_plot_options(column):
    with column:
        st.header("Plot options:")
        ti = st.text_input(
            "Enter title for the graph:"
        )
        xaxis_label = st.text_input(
            "Enter name for x-axis:",
        )
        yaxis_label = st.text_input(
            "Enter name for y-axis:",
        )
        bool_sort_values = st.checkbox("Sort values for plot")
    return ti, xaxis_label, yaxis_label, bool_sort_values

st.set_page_config(
    page_title = "Excel-Plotter",
    page_icon="📈",
    layout = 'wide',
)

st.title('Data analysis of Excel documents 📈')
st.markdown('---')

excel_file = st.file_uploader('Input the Excel-file:', type='xlsx', accept_multiple_files=False)
st.markdown('---')
c1, c2 = st.columns((2,1))


if excel_file:
    excel_pd = pd.read_excel(excel_file, engine = 'openpyxl')
    columns = list(excel_pd.columns)
    st.sidebar.header("Data options: ")
    column_chosen = st.sidebar.selectbox(
        "Select column to analyze: ",
        options=columns
    )

    bokning = excel_pd[column_chosen].dropna().tolist()
    dict = {}
    dict_occurrences = {bokning[i]: int(bokning.count(bokning[i])) for i in range(len(bokning))}
    dict_percentage = {bokning[i]: (x := int(bokning.count(bokning[i])),1,x) for i in range(len(bokning))}
    calc_radio = st.sidebar.radio(
        "Choose how you want to study the data",
        ('Calculate and plot number of occurrences in total for each entry', 'Calculate and plot percentages for each entry')
        )
    ti, xaxis_label, yaxis_label, bool_sort_values = display_plot_options(c2)
    if calc_radio == 'Calculate and plot percentages for each entry':
        dict = dict_percentage
        df_bokning = pd.DataFrame(dict).T.rename_axis(column_chosen).reset_index()
        df_bokning.columns = [column_chosen,'Number of occurrences', 'Max number of occurrences', 'Percentage']
        with st.sidebar.form(f'{dict}'):
            st.write("Input max number of occurrences for each entry. Press submit after.")
            for key in dict.keys():
                temp_val = dict_occurrences.get(key)
                val = st.number_input(f'{key}:', min_value = 1)
                dict[key] = (temp_val, val, temp_val/val)       
            submitted = st.form_submit_button("Submit")
            if submitted:
                df_bokning = pd.DataFrame(dict).T.rename_axis(column_chosen).reset_index()
                df_bokning.columns = [column_chosen,'Number of occurrences', 'Max number of occurrences', 'Percentage']
                
        bool_sort_values = c2.checkbox("Sort") 
        if bool_sort_values:
            df_bokning = df_bokning.sort_values(df_bokning.columns[-1])
        generate_plot(df_bokning, c1, xaxis_label, yaxis_label, ti)
        st.markdown('---')
        st.header('Excel-table of data')
        st.dataframe(df_bokning)
        generate_excel_download_link(df_bokning)
    else:
        dict = dict_occurrences
        df_bokning = pd.DataFrame(dict, index=[0]).T.rename_axis(column_chosen).reset_index()
        df_bokning.columns = [column_chosen,'Number of occurrences']
        if bool_sort_values:
            df_bokning = df_bokning.sort_values(df_bokning.columns[-1]) 
        generate_plot(df_bokning, c1, xaxis_label, yaxis_label, ti)
        st.markdown('---')
        st.header('Excel-table of data')
        st.dataframe(df_bokning)
        generate_excel_download_link(df_bokning)

