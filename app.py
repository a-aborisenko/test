try:
    import openpyxl
except ImportError:
    st.error("openpyxl не установлен.")
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Генератор отчётов по времени", page_icon="⏱️", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #2b2b2b; color: #f0f0f0; }
    h1, h2, h3, h4 { color: #9b59b6 !important; }
    .stButton>button, .stDownloadButton>button {
        background-color: #9b59b6; color: #ffffff; border: none; border-radius: 8px; padding: 0.6rem 1rem;
    }
    .stButton>button:hover, .stDownloadButton>button:hover { background-color: #884ea0; color: #ffffff; }
    section[data-testid="stSidebar"] { background-color: #1e1e1e; }
    </style>
""", unsafe_allow_html=True)

st.title("Генератор отчётов по времени")
st.write("Загрузите Excel-файл (.xlsx) для анализа")

uploaded_file = st.file_uploader("Перетащите или выберите Excel-файл", type=["xlsx"])

def process_timesheet(df):
    if not all(col in df.columns for col in ['Имя активности', 'Полное название', 'Записанные часы']):
        st.error("Файл должен содержать столбцы: 'Имя активности', 'Полное название', 'Записанные часы'")
        return None
    df['Записанные часы'] = pd.to_numeric(df['Записанные часы'], errors='coerce').fillna(0)
    result = df.groupby(['Имя активности','Полное название'])['Записанные часы'].sum().reset_index()
    result.columns = ['Проект','Специалист','Часы']
    return result

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.success("Файл загружен!")
        result = process_timesheet(df)
        if result is not None:
            st.subheader("Предварительный просмотр")
            st.dataframe(result.head(10))
    except Exception as e:
        st.error(f"Ошибка: {e}")
