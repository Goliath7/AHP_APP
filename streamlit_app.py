import datetime
import json
from pathlib import Path

import streamlit as st

from letter_doc_builder import format_dates_for_shift, generate_letter_document
from nariad_gui_final import generate_naryad_document
from plan_rabot_GUI import generate_document as generate_plan_document
from shared_data import LEADERS_FULL, LETTER_SUPERVISORS, STATIONS

APP_DIR = Path(__file__).resolve().parent
ADMIN_DATA_FILE = APP_DIR / "admin_data.json"


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


def _normalize_list(items):
    result = []
    seen = set()
    for item in items:
        value = str(item).strip()
        if value and value not in seen:
            result.append(value)
            seen.add(value)
    return result


def _normalize_supervisors(items):
    result = []
    seen = set()
    for name, phone in items:
        n = str(name).strip()
        p = str(phone).strip()
        if n and (n, p) not in seen:
            result.append((n, p))
            seen.add((n, p))
    return result


def _load_admin_data_from_file():
    if not ADMIN_DATA_FILE.exists():
        return None
    try:
        raw = json.loads(ADMIN_DATA_FILE.read_text(encoding="utf-8"))
        stations = _normalize_list(raw.get("stations", []))
        leaders = _normalize_list(raw.get("leaders", []))
        supervisors_raw = raw.get("letter_supervisors", [])
        supervisors = []
        for item in supervisors_raw:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                supervisors.append((item[0], item[1]))
        supervisors = _normalize_supervisors(supervisors)
        if stations and leaders and supervisors:
            return {
                "stations": stations,
                "leaders": leaders,
                "letter_supervisors": supervisors,
            }
    except Exception:
        return None
    return None


def _save_admin_data_to_file(stations, leaders, letter_supervisors):
    payload = {
        "stations": stations,
        "leaders": leaders,
        "letter_supervisors": letter_supervisors,
    }
    ADMIN_DATA_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _init_reference_data():
    if "cfg_stations" in st.session_state:
        return

    loaded = _load_admin_data_from_file()
    if loaded:
        st.session_state.cfg_stations = loaded["stations"]
        st.session_state.cfg_leaders = loaded["leaders"]
        st.session_state.cfg_letter_supervisors = loaded["letter_supervisors"]
    else:
        st.session_state.cfg_stations = list(STATIONS)
        st.session_state.cfg_leaders = list(LEADERS_FULL)
        st.session_state.cfg_letter_supervisors = list(LETTER_SUPERVISORS)


def _get_stations():
    return st.session_state.get("cfg_stations", list(STATIONS))


def _get_leaders():
    return st.session_state.get("cfg_leaders", list(LEADERS_FULL))


def _get_letter_supervisors():
    return st.session_state.get("cfg_letter_supervisors", list(LETTER_SUPERVISORS))


def _admin_password():
    try:
        return st.secrets.get("admin_password", "admin123")
    except Exception:
        return "admin123"


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
    stations = _get_stations()
    letter_supervisors = _get_letter_supervisors()

    if not stations:
        st.error("Список станций пуст. Заполните его во вкладке 'Админ'.")
        return
    if not letter_supervisors:
        st.error("Список руководителей с телефонами пуст. Заполните его во вкладке 'Админ'.")
        return

    col1, col2 = st.columns([1, 2])
    with col1:
        num = st.text_input("Исх. №", value="263")
    with col2:
        contract = st.text_input("Договор", value="№ 4313м от 10.06.2024 г.")

    if "letter_items" not in st.session_state:
        st.session_state.letter_items = []

    with st.form("letter_item_form", clear_on_submit=False):
        station = st.selectbox("Станция", stations, key="letter_station")
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
            [f"{name} — {phone}" for name, phone in letter_supervisors],
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
    stations = _get_stations()
    leaders = _get_leaders()
    if not stations or not leaders:
        st.error("Справочники станций/руководителей пусты. Заполните их во вкладке 'Админ'.")
        return

    station = st.selectbox("Станция", stations, key="naryad_station")
    leader = st.selectbox("Руководитель работ", leaders, key="naryad_leader")

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
    stations = _get_stations()
    leaders = _get_leaders()
    if not stations or not leaders:
        st.error("Справочники станций/руководителей пусты. Заполните их во вкладке 'Админ'.")
        return

    station = st.selectbox("Станция", stations, key="plan_station")
    supervisor = st.selectbox("Руководитель", leaders, key="plan_supervisor")

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


def _admin_tab():
    st.subheader("Админ-панель")

    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False

    if not st.session_state.admin_authenticated:
        st.info("Вход в админ-панель")
        password = st.text_input("Пароль администратора", type="password", key="admin_password_input")
        col1, col2 = st.columns([1, 5])
        with col1:
            enter = st.button("Войти")
        with col2:
            if _admin_password() == "admin123":
                st.warning("Используется пароль по умолчанию `admin123`. Лучше задать `admin_password` в секретах Streamlit.")
        if enter:
            if password == _admin_password():
                st.session_state.admin_authenticated = True
                st.success("Доступ разрешён")
                st.rerun()
            else:
                st.error("Неверный пароль")
        return

    cols = st.columns([1, 1, 4])
    with cols[0]:
        if st.button("Выйти"):
            st.session_state.admin_authenticated = False
            st.rerun()
    with cols[1]:
        if st.button("Сбросить формы"):
            st.session_state.letter_items = []
            st.success("Формы очищены")

    stations_text = st.text_area(
        "Станции (по одной на строку)",
        value="\n".join(_get_stations()),
        height=220,
    )
    leaders_text = st.text_area(
        "Руководители (ФИО, по одному на строку)",
        value="\n".join(_get_leaders()),
        height=220,
    )
    supervisors_text = st.text_area(
        "Руководители для письма с телефонами (формат: ФИО | телефон)",
        value="\n".join([f"{name} | {phone}" for name, phone in _get_letter_supervisors()]),
        height=220,
    )

    if st.button("Сохранить справочники", type="primary"):
        try:
            stations = _normalize_list(stations_text.splitlines())
            leaders = _normalize_list(leaders_text.splitlines())
            supervisors = []
            for line in supervisors_text.splitlines():
                value = line.strip()
                if not value:
                    continue
                if "|" in value:
                    name, phone = [part.strip() for part in value.split("|", 1)]
                else:
                    name, phone = value, "не указан"
                supervisors.append((name, phone))
            supervisors = _normalize_supervisors(supervisors)

            if not stations:
                raise ValueError("Нужно указать хотя бы одну станцию")
            if not leaders:
                raise ValueError("Нужно указать хотя бы одного руководителя")
            if not supervisors:
                raise ValueError("Нужно указать хотя бы одного руководителя с телефоном")

            st.session_state.cfg_stations = stations
            st.session_state.cfg_leaders = leaders
            st.session_state.cfg_letter_supervisors = supervisors
            _save_admin_data_to_file(stations, leaders, supervisors)
            st.success(f"Справочники сохранены ({ADMIN_DATA_FILE.name})")
        except Exception as error:
            st.error(str(error))


def main():
    st.set_page_config(page_title="Генератор документов", page_icon="📄", layout="wide")
    _inject_styles()
    _init_reference_data()

    st.markdown(
        """
        <div class="title-box">
          <h1>Единый генератор документов</h1>
          <p>Исходящее письмо, наряд и план работ в веб-интерфейсе</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tab1, tab2, tab3, tab4 = st.tabs(["Исходящее письмо", "Наряд", "План работ", "Админ"])
    with tab1:
        _letter_tab()
    with tab2:
        _naryad_tab()
    with tab3:
        _plan_tab()
    with tab4:
        _admin_tab()


if __name__ == "__main__":
    main()
