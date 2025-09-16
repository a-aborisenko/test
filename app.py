import streamlit as st
import pandas as pd
import openpyxl
import io

# --- Функция обработки данных ---
def process_timesheet(df, proj_filter=None):
    required_cols = {'Имя активности', 'Полное название', 'Записанные часы'}
    if not required_cols.issubset(set(df.columns)):
        missing = required_cols - set(df.columns)
        raise ValueError(f"В файле отсутствуют необходимые столбцы: {', '.join(missing)}")
    if df.empty:
        raise ValueError("Файл пуст или не содержит данных.")
    # Проверка числовых значений
    df['Записанные часы'] = pd.to_numeric(df['Записанные часы'], errors='coerce')
    if df['Записанные часы'].isnull().any():
        raise ValueError("В столбце 'Записанные часы' есть нечисловые значения.")
    # Фильтрация по проекту (если требуется)
    if proj_filter:
        df = df[df['Имя активности'] == proj_filter]
    # Группировка и округление
    result = (
        df.groupby(['Имя активности', 'Полное название'])['Записанные часы']
        .sum()
        .reset_index()
        .sort_values(['Имя активности', 'Полное название'])
    )
    result['Записанные часы'] = result['Записанные часы'].round(2)
    result.columns = ['Проект', 'Специалист', 'Часы']
    return result

# --- Функция создания ссылки для скачивания ---
def create_download_link(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- Интерфейс Streamlit ---
def main():
    st.set_page_config(page_title="Генератор отчётов по времени", layout="wide")
    st.markdown(
        "<h1 style='color:#1f77b4;'>Генератор отчётов по времени</h1>", unsafe_allow_html=True)
    st.write("Загрузите Excel-файл (.xlsx) с данными учёта времени. После обработки появится сводная таблица и статистика.")
    uploaded_file = st.file_uploader("Drag-and-drop Excel (.xlsx)", type="xlsx")
    proj_filter = None
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            # Показываем фильтр по проекту, если есть проекты
            all_projects = sorted(df['Имя активности'].unique())
            proj_filter = st.selectbox("Фильтр по проекту", ["Все"] + all_projects)
            project = proj_filter if proj_filter != "Все" else None
            with st.spinner("Обработка данных..."):
                result = process_timesheet(df, project)
            # Статистика
            n_projects = result['Проект'].nunique()
            n_specialists = result['Специалист'].nunique()
            total_hours = result['Часы'].sum()
            st.success("Данные успешно обработаны!")
            st.write(f"**Уникальных проектов:** {n_projects}")
            st.write(f"**Уникальных специалистов:** {n_specialists}")
            st.write(f"**Всего часов:** {total_hours:.2f}")
            # Просмотр таблицы
            st.subheader("Первые 10 строк отчёта")
            st.dataframe(result.head(10))
            # Кнопка скачивания
            st.download_button("Скачать отчёт (Excel)", data=create_download_link(result), file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Ошибка: {str(e)}")
    else:
        st.info("Загрузите .xlsx файл для обработки.")

if __name__ == "__main__":
    main()
