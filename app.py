import streamlit as st
import pandas as pd

# ================== СТИЛИ ==================
st.markdown("""
    <style>
    /* Общий фон */
    .stApp {
        background-color: #2b2b2b;
        color: #f0f0f0;
    }

    /* Заголовки */
    h1, h2, h3, h4 {
        color: #9b59b6 !important;
    }

    /* Текст */
    .stMarkdown p, .stText, .stDataFrame {
        color: #f0f0f0 !important;
    }

    /* Кнопки */
    .stButton>button, .stDownloadButton>button {
        background-color: #9b59b6;
        color: #ffffff;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1rem;
    }

    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #884ea0;
        color: #ffffff;
    }

    /* Сайдбар */
    section[data-testid="stSidebar"] {
        background-color: #1e1e1e;
    }
    </style>
""", unsafe_allow_html=True)

# ================== ПРОВЕРКА ЗАВИСИМОСТЕЙ ==================
st.sidebar.header("Проверка зависимостей")
try:
    import openpyxl
    st.sidebar.success("✅ Все зависимости установлены: pandas, openpyxl")
except ImportError as e:
    st.sidebar.error(f"❌ Проблема с зависимостями: {e}")

# ================== ГЛАВНЫЙ ЭКРАН ==================
st.title("Генератор отчётов по времени")
st.write("Загрузите Excel-файл и получите сводный отчёт по проектам и специалистам.")

# ================== ЗАГРУЗКА ФАЙЛА ==================
uploaded_file = st.file_uploader("Перетащите файл .xlsx или выберите его", type=["xlsx"])

def process_timesheet(df):
    if not all(col in df.columns for col in ['Имя активности', 'Полное название', 'Записанные часы']):
        st.error("Файл должен содержать столбцы: 'Имя активности', 'Полное название', 'Записанные часы'")
        return None

    result = df.groupby(['Имя активности', 'Полное название'])['Записанные часы'].sum().reset_index()
    result.columns = ['Проект', 'Специалист', 'Часы']
    return result

# ================== ОБРАБОТКА И ВЫГРУЗКА ==================
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.success("Файл успешно загружен!")
        result = process_timesheet(df)
        if result is not None:
            st.subheader("Предварительный просмотр")
            st.dataframe(result.head(10))

            # Скачивание отчёта
            @st.cache_data
            def convert_df_to_excel(dataframe):
                return dataframe.to_excel(index=False, engine='openpyxl')

            excel_data = convert_df_to_excel(result)
            st.download_button(
                label="Скачать отчёт в Excel",
                data=excel_data,
                file_name="report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
