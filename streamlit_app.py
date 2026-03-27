import datetime
from pathlib import Path

import streamlit as st

from letter_doc_builder import format_dates_for_shift, generate_letter_document
from nariad_gui_final import generate_naryad_document
from plan_rabot_GUI import generate_document as generate_plan_document
from shared_data import LEADERS_FULL, LETTER_SUPERVISORS, STATIONS


def _inject_styles():
    st.markdown(
        """
        <style>
        .main {
            background: linear-gradient(180deg, #f4f7fb 0%, #eef3f9 100%);
        }
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1100px;
        }
        .title-box {
            padding: 1rem 1.2rem;
            border-radius: 12px;
            background: linear-gradient(135deg, #14395b 0%, #1f5f8f 100%);
            color: white;
            margin-bottom: 1rem;
        }
        .title-box h1 {
            margin: 0;
            font-size: 1.6rem;
            font-weight: 700;
        }
        .title-box p {
            margin: 0.35rem 0 0;
            color: #d8ecff;
            font-size: 0.95rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _download_file(path, label, key):
    file_path = Path(path)
    with file_path.open("rb") as file_stream:
        st.download_button(
            label=label,
            data=file_stream.read(),
            file_name=file_path.name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=key,
        )


def _expand_dates(mode, single_date, range_dates):
    if mode == "Одна дата":
        return [single_date]

    if not isinstance(range_dates, tuple) or len(range_dates) != 2:
        raise ValueError("Выберите период из двух дат")

    start_date, end_date = range_dates
    if start_date > end_date:
        raise ValueError("Дата начала периода позже даты окончания")

    delta_days = (end_date - start_date).days
    return [start_date + datetime.timedelta(days=i) for i in range(delta_days + 1)]


def _letter_tab():
    st.subheader("Исходящее письмо")

    col1, col2 = st.columns([1, 2])
    with col1:
        num = st.text_input("Исх. №", value="263")
    with col2:
        contract = st.text_input("Договор", value="№ 4313м от 10.06.2024 г.")

    if "letter_items" not in st.session_state:
        st.session_state.letter_items = []

    with st.form("letter_item_form", clear_on_submit=False):
        station = st.selectbox("Станция", STATIONS, key="letter_station")
        shift = st.radio("Смена", ["Дневная", "Ночная"], horizontal=True, key="letter_shift")

        date_mode = st.radio("Режим дат", ["Одна дата", "Период"], horizontal=True, key="letter_date_mode")
        today = datetime.date.today()
        if date_mode == "Одна дата":
            single_date = st.date_input("Дата", value=today, key="letter_single_date")
            range_dates = None
        else:
            single_date = today
            range_dates = st.date_input(
                "Период дат",
                value=(today, today + datetime.timedelta(days=1)),
                key="letter_range_date",
            )

        sup_option = st.selectbox(
            "Руководитель работ",
            [f"{name} — {phone}" for name, phone in LETTER_SUPERVISORS],
            key="letter_supervisor",
        )

        add_item = st.form_submit_button("Добавить позицию")

    if add_item:
        try:
            dates = _expand_dates(date_mode, single_date, range_dates)
            dates_str = format_dates_for_shift(dates, shift)
            sup_name, sup_phone = [item.strip() for item in sup_option.split("—", 1)]
            st.session_state.letter_items.append(
                {
                    "station": station,
                    "dates_str": dates_str,
                    "sup_name": sup_name,
                    "sup_phone": sup_phone,
                }
            )
            st.success("Позиция добавлена")
        except Exception as error:
            st.error(str(error))

    if st.session_state.letter_items:
        st.markdown("**Текущий список позиций:**")
        st.dataframe(
            [
                {
                    "Станция": item["station"],
                    "Даты/смена": item["dates_str"],
                    "Руководитель": item["sup_name"],
                    "Телефон": item["sup_phone"],
                }
                for item in st.session_state.letter_items
            ],
            use_container_width=True,
            hide_index=True,
        )

        remove_idx = st.number_input(
            "Удалить позицию №",
            min_value=1,
            max_value=len(st.session_state.letter_items),
            value=1,
            step=1,
        )
        if st.button("Удалить выбранную позицию"):
            del st.session_state.letter_items[int(remove_idx) - 1]
            st.rerun()

    if st.button("Сформировать исходящее письмо", type="primary"):
        try:
            work_list = [
                (item["station"], item["dates_str"], item["sup_name"], item["sup_phone"])
                for item in st.session_state.letter_items
            ]
            file_path = generate_letter_document(num=num, contract=contract, work_list=work_list)
            st.success(f"Документ создан: {file_path}")
            _download_file(file_path, "Скачать исходящее письмо", "download_letter")
        except Exception as error:
            st.error(str(error))


def _naryad_tab():
    st.subheader("Наряд")
    station = st.selectbox("Станция", STATIONS, key="naryad_station")
    leader = st.selectbox("Руководитель работ", LEADERS_FULL, key="naryad_leader")

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Начало работ", value=datetime.date.today(), key="naryad_start")
    with col2:
        end_date = st.date_input("Окончание работ", value=datetime.date.today() + datetime.timedelta(days=10), key="naryad_end")

    if st.button("Сформировать наряд", type="primary"):
        try:
            file_path = generate_naryad_document(station, leader, start_date, end_date)
            st.success(f"Документ создан: {file_path}")
            _download_file(file_path, "Скачать наряд", "download_naryad")
        except Exception as error:
            st.error(str(error))


def _plan_tab():
    st.subheader("План работ")
    station = st.selectbox("Станция", STATIONS, key="plan_station")
    supervisor = st.selectbox("Руководитель", LEADERS_FULL, key="plan_supervisor")

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Начало работ", value=datetime.date.today(), key="plan_start")
    with col2:
        end_date = st.date_input("Окончание работ", value=datetime.date.today() + datetime.timedelta(days=10), key="plan_end")

    if st.button("Сформировать план работ", type="primary"):
        try:
            file_path = generate_plan_document(
                station=station,
                supervisor_full=supervisor,
                start_dt=start_date,
                end_dt=end_date,
                show_messages=False,
            )
            st.success(f"Документ создан: {file_path}")
            _download_file(file_path, "Скачать план работ", "download_plan")
        except Exception as error:
            st.error(str(error))


def main():
    st.set_page_config(page_title="Генератор документов", page_icon="📄", layout="wide")
    _inject_styles()

    st.markdown(
        """
        <div class="title-box">
          <h1>Единый генератор документов</h1>
          <p>Исходящее письмо, наряд и план работ в веб-интерфейсе</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3 = st.tabs(["Исходящее письмо", "Наряд", "План работ"])
    with tab1:
        _letter_tab()
    with tab2:
        _naryad_tab()
    with tab3:
        _plan_tab()


if __name__ == "__main__":
    main()
