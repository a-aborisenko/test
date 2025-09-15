# -*- coding: utf-8 -*-
"""
Генератор отчётов по времени — Streamlit-приложение
Автор: вы :)
Назначение: загрузка .xlsx табелей, группировка часов по проектам и специалистам,
формирование сводного отчёта и выгрузка в Excel.

Технологии: Streamlit, pandas, openpyxl
Хостинг: Streamlit Community Cloud (бесплатно)
"""

import io
import logging
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ----------------------------- НАСТРОЙКИ СТРАНИЦЫ -----------------------------

st.set_page_config(
    page_title="Генератор отчётов по времени",
    page_icon="⏱️",
    layout="wide",
)

# Минималистичный flat design + акценты
PRIMARY = "#1f77b4"  # синие акценты
TEXT_GRAY = "#666666"  # серый текст

st.markdown(
    f"""
    <style>
      html, body, [class*="css"]  {{
        color: {TEXT_GRAY};
        background: #ffffff;
      }}
      .stApp h1, .stApp h2, .stApp h3 {{
        color: {PRIMARY};
        font-weight: 700;
      }}
      .stButton>button, .stDownloadButton>button {{
        background: {PRIMARY};
        color: #fff;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1rem;
      }}
      .stButton>button:hover, .stDownloadButton>button:hover {{
        background: #15609a;
      }}
      .stProgress>div>div>div>div {{
        background-color: {PRIMARY};
      }}
      /* Убираем тени, делаем flat */
      .stCard, .stDataFrame, .block-container {{
        box-shadow: none !important;
      }}
      /* Аккуратные контейнеры */
      .metric-container {{
        background: #f8f9fb;
        border: 1px solid #eef0f4;
        border-radius: 12px;
        padding: 12px;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------- ЛОГИРОВАНИЕ ----------------------------------

logger = logging.getLogger("timesheet_app")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler()
    formatter = logging.Formatter("%(asctime)s — %(levelname)s — %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

if "logs" not in st.session_state:
    st.session_state.logs = []


def log(msg: str, level: str = "info"):
    """Пишем в системный лог и в UI-лог"""
    if level == "error":
        logger.error(msg)
    elif level == "warning":
        logger.warning(msg)
    else:
        logger.info(msg)
    st.session_state.logs.append(msg)


# ---------------------------- СЛУЖЕБНЫЕ ФУНКЦИИ -------------------------------

# Допустимые "алиасы" названий столбцов (легко расширить)
ACTIVITY_ALIASES = {"имя активности", "проект", "activity name", "activity"}
PERSON_ALIASES = {"полное название", "сотрудник", "специалист", "full name", "employee"}
HOURS_ALIASES = {"записанные часы", "часы", "hours", "logged hours", "time"}


def normalize(s: str) -> str:
    return str(s).strip().lower().replace("\n", " ").replace("\r", " ")


def find_column_by_alias(
    df: pd.DataFrame, aliases: set
) -> Optional[str]:
    """Ищем столбец по набору алиасов (по нормализованным заголовкам)."""
    name_map = {normalize(c): c for c in df.columns}
    for key_norm, orig in name_map.items():
        if key_norm in aliases:
            return orig
    return None


def fallback_by_excel_letters(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Фолбэк на позиции столбцов по буквам Excel:
    V (22-я), G (7-я), C (3-я).
    Возвращает словарь {role: column_name_or_None}.
    """
    pos_map = {"activity": 21, "person": 6, "hours": 2}  # 0-индексация: V=21, G=6, C=2
    res = {}
    for role, idx in pos_map.items():
        if idx < len(df.columns):
            res[role] = df.columns[idx]
        else:
            res[role] = None
    return res


def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Пытаемся автоматически определить столбцы: проект, специалист, часы.
    Сначала ищем по алиасам, затем — фолбэком по буквам.
    """
    activity = find_column_by_alias(df, ACTIVITY_ALIASES)
    person = find_column_by_alias(df, PERSON_ALIASES)
    hours = find_column_by_alias(df, HOURS_ALIASES)

    auto = {"activity": activity, "person": person, "hours": hours}

    if not all(auto.values()):
        fb = fallback_by_excel_letters(df)
        auto = {k: (auto[k] or fb[k]) for k in auto}

    return auto


def validate_hours(series: pd.Series) -> Tuple[pd.Series, int]:
    """
    Приводим значения часов к числу (float).
    Возвращаем очищенную серию и количество некорректных значений.
    """
    coerced = pd.to_numeric(series, errors="coerce")
    invalid_count = int(coerced.isna().sum() - series.isna().sum())
    # NaN из-за некорректных строк => считаем как 0 при агрегации,
    # но показываем предупреждение пользователю.
    coerced = coerced.fillna(0)
    return coerced, invalid_count


def process_timesheet(
    df: pd.DataFrame,
    cols: Dict[str, str],
) -> Tuple[pd.DataFrame, Dict[str, float]]:
    """
    Основная обработка данных:
    - проверка столбцов
    - валидация часов
    - группировка по проектам и специалистам
    - сортировка и форматирование
    Возвращает (итоговая_таблица, статистика).
    """
    required = ["activity", "person", "hours"]
    for r in required:
        if r not in cols or cols[r] not in df.columns:
            raise ValueError(
                "Не найдены требуемые столбцы. Проверьте, что доступны V/\"Имя активности\", "
                "G/\"Полное название\", C/\"Записанные часы\" — либо выберите вручную в боковой панели."
            )

    work = df[[cols["activity"], cols["person"], cols["hours"]]].copy()
    work.columns = ["Проект", "Специалист", "Часы"]

    # Пустые строки удаляем
    work = work.dropna(subset=["Проект", "Специалист"], how="any")

    # Валидация часов
    work["Часы"], invalid_count = validate_hours(work["Часы"])

    # Группировка
    grouped = (
        work.groupby(["Проект", "Специалист"], as_index=False)["Часы"].sum()
    )

    # Итоговая сортировка
    grouped = grouped.sort_values(["Проект", "Специалист"], kind="mergesort").reset_index(drop=True)

    # Округление (для экспорта как число с 2 знаками)
    grouped["Часы"] = grouped["Часы"].round(2)

    # Статистика
    stats = {
        "projects": float(grouped["Проект"].nunique()),
        "people": float(grouped["Специалист"].nunique()),
        "hours_total": float(grouped["Часы"].sum()),
        "invalid_rows": float(invalid_count),
        "source_rows": float(len(df)),
        "used_rows": float(len(work)),
    }

    return grouped, stats


def format_preview(df: pd.DataFrame, limit: int = 10) -> pd.DataFrame:
    """Форматирование превью (часы — с двумя знаками после запятой)."""
    prev = df.head(limit).copy()
    prev["Часы"] = prev["Часы"].map(lambda x: f"{x:.2f}")
    return prev


def create_excel_report(
    df: pd.DataFrame,
    stats: Dict[str, float],
    sheet_name_report: str = "Отчёт",
    sheet_name_stats: str = "Статистика",
) -> bytes:
    """
    Создаёт Excel в памяти:
    - Лист "Отчёт" (Проект, Специалист, Часы с форматом 0.00)
    - Лист "Статистика"
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name_report)
        stats_df = pd.DataFrame(
            {
                "Показатель": ["Количество проектов", "Количество специалистов", "Всего часов",
                               "Некорректных часов (приведены к 0)", "Строк в исходных данных", "Строк после очистки"],
                "Значение": [int(stats["projects"]), int(stats["people"]), stats["hours_total"],
                             int(stats["invalid_rows"]), int(stats["source_rows"]), int(stats["used_rows"])],
            }
        )
        stats_df.to_excel(writer, index=False, sheet_name=sheet_name_stats)

        # Форматируем столбец "Часы" как 0.00
        wb = writer.book
        ws_report = writer.sheets[sheet_name_report]

        # Поиск индекса столбца "Часы"
        hours_col_idx = list(df.columns).index("Часы") + 1  # openpyxl 1-индексация
        for row in range(2, len(df) + 2):  # со 2-й строки (после заголовка)
            cell = ws_report.cell(row=row, column=hours_col_idx)
            cell.number_format = "0.00"

        # Заморозка верхней строки
        ws_report.freeze_panes = "A2"

        # Автоширина столбцов
        for ws in [ws_report, writer.sheets[sheet_name_stats]]:
            for column_cells in ws.columns:
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max(10, max_len + 2), 60)

    return output.getvalue()


# ---------------------------------- UI ----------------------------------------

st.title("Генератор отчётов по времени")
st.write("Загрузите Excel-файл (.xlsx) с данными учёта времени и получите сводный отчёт по проектам и специалистам.")

with st.expander("🛈 Подсказка по столбцам", expanded=False):
    st.markdown(
        """
        Приложение ожидает:
        - **V «Имя активности»** — названия проектов  
        - **G «Полное название»** — имена специалистов  
        - **C «Записанные часы»** — количество часов  

        Если заголовки отличаются, используйте **боковую панель** справа, чтобы выбрать нужные столбцы вручную.
        """
    )

file = st.file_uploader("Перетащите файл .xlsx или выберите вручную", type=["xlsx"])

# Боковая панель — ручной выбор столбцов
st.sidebar.header("Настройки обработки")
st.sidebar.caption("При необходимости выберите столбцы вручную")
manual_cols = {"activity": None, "person": None, "hours": None}

process_clicked = False

if file is not None:
    try:
        log("Чтение загруженного файла...")
        df_raw = pd.read_excel(file, engine="openpyxl")

        st.sidebar.subheader("Выбор столбцов")
        # Список столбцов файла
        cols_list = list(df_raw.columns)

        # Автоопределение
        auto = detect_columns(df_raw)

        activity_col = st.sidebar.selectbox(
            "Столбец проекта",
            options=["(Авто)"] + cols_list,
            index=0 if auto["activity"] is None else (cols_list.index(auto["activity"]) + 1),
        )
        person_col = st.sidebar.selectbox(
            "Столбец специалиста",
            options=["(Авто)"] + cols_list,
            index=0 if auto["person"] is None else (cols_list.index(auto["person"]) + 1),
        )
        hours_col = st.sidebar.selectbox(
            "Столбец часов",
            options=["(Авто)"] + cols_list,
            index=0 if auto["hours"] is None else (cols_list.index(auto["hours"]) + 1),
        )

        manual_cols["activity"] = None if activity_col == "(Авто)" else activity_col
        manual_cols["person"] = None if person_col == "(Авто)" else person_col
        manual_cols["hours"] = None if hours_col == "(Авто)" else hours_col

        st.sidebar.markdown("---")
        st.sidebar.subheader("Фильтр проекта (после обработки)")
        st.sidebar.caption("Станет доступен после нажатия «Обработать данные»")

    except Exception as e:
        st.error(f"Ошибка чтения Excel: {e}")
        log(f"Ошибка чтения Excel: {e}", level="error")

# Кнопка обработки
col_btn, _ = st.columns([1, 3])
with col_btn:
    process_clicked = st.button("Обработать данные", use_container_width=True)

# --------------------------- ПРОЦЕСС ОБРАБОТКИ --------------------------------

if process_clicked:
    if file is None:
        st.warning("Сначала загрузите .xlsx файл.")
    else:
        progress = st.progress(0, text="Начало обработки...")
        try:
            # 1) Проверка формата
            if not str(file.name).lower().endswith(".xlsx"):
                raise ValueError("Неверный формат. Поддерживаются только файлы .xlsx")

            progress.progress(15, text="Проверка столбцов...")
            # 2) Определяем столбцы (ручной выбор приоритетнее)
            auto = detect_columns(df_raw)
            cols = {
                "activity": manual_cols["activity"] or auto["activity"],
                "person": manual_cols["person"] or auto["person"],
                "hours": manual_cols["hours"] or auto["hours"],
            }

            # 3) Базовые проверки
            if df_raw.empty:
                raise ValueError("Файл пуст или не содержит данных.")
            missing = [k for k, v in cols.items() if v is None]
            if missing:
                raise ValueError(
                    "Не удалось определить требуемые столбцы. "
                    "Укажите их в боковой панели или проверьте заголовки."
                )

            progress.progress(45, text="Группировка и суммирование...")
            result_df, stats = process_timesheet(df_raw, cols)

            progress.progress(70, text="Подготовка превью и статистики...")

            # Фильтр по проекту (сайдбар)
            projects = ["Все проекты"] + sorted(result_df["Проект"].unique().tolist())
            selected_project = st.sidebar.selectbox("Проект", options=projects, index=0)

            if selected_project != "Все проекты":
                filtered_df = result_df[result_df["Проект"] == selected_project].reset_index(drop=True)
            else:
                filtered_df = result_df

            # Метрики
            mcol1, mcol2, mcol3 = st.columns(3)
            with mcol1:
                st.markdown('<div class="metric-container">', unsafe_allow_html=True)
                st.metric("Проектов", int(stats["projects"]))
                st.markdown("</div>", unsafe_allow_html=True)
            with mcol2:
                st.markdown('<div class="metric-container">', unsafe_allow_html=True)
                st.metric("Специалистов", int(stats["people"]))
                st.markdown("</div>", unsafe_allow_html=True)
            with mcol3:
                st.markdown('<div class="metric-container">', unsafe_allow_html=True)
                st.metric("Всего часов", f'{stats["hours_total"]:.2f}')
                st.markdown("</div>", unsafe_allow_html=True)

            if stats["invalid_rows"] > 0:
                st.info(
                    f"Обнаружено некорректных значений в столбце «Часы»: {int(stats['invalid_rows'])}. "
                    "Они были автоматически приведены к 0."
                )

            # Превью (первые 10 строк)
            st.subheader("Предварительный просмотр (первые 10 строк)")
            st.dataframe(format_preview(filtered_df, limit=10), use_container_width=True, hide_index=True)

            progress.progress(85, text="Формирование Excel-отчёта...")

            # Экспорт
            excel_bytes = create_excel_report(filtered_df, stats)
            file_name = "timesheet_report.xlsx"
            st.download_button(
                label="Скачать отчёт (Excel)",
                data=excel_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            progress.progress(100, text="Готово!")
            log("Обработка успешно завершена.")

            # Показываем журнал операций
            with st.expander("Журнал операций"):
                for line in st.session_state.logs:
                    st.write("• " + line)

        except Exception as e:
            progress.empty()
            st.error(f"Ошибка обработки: {e}")
            log(f"Ошибка обработки: {e}", level="error")


# -------------------------- КОНЕЦ ПРИЛОЖЕНИЯ ----------------------------------
