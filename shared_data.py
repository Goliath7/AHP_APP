from pathlib import Path


def _dedupe_keep_order(items):
    seen = set()
    result = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


def fio_short(full_name):
    full_name = " ".join(full_name.strip().split())
    if not full_name:
        return ""

    # Уже в кратком формате: "Фамилия И.О." или похожее.
    if "." in full_name:
        return full_name

    parts = full_name.split()
    if len(parts) >= 3:
        return f"{parts[0]} {parts[1][0]}. {parts[2][0]}."
    if len(parts) == 2:
        return f"{parts[0]} {parts[1][0]}."
    return parts[0]


def safe_filename(name):
    forbidden = '<>:"/\\|?*'
    cleaned = name.strip()
    for ch in forbidden:
        cleaned = cleaned.replace(ch, "_")
    return cleaned


BASE_DIR = Path(__file__).resolve().parent


def output_path(filename):
    return BASE_DIR / safe_filename(filename)


STATIONS = _dedupe_keep_order([
    "Смоленская АПЛ", "Арбатская ФЛ", "Курская КЛ", "Новокузнецкая", "Смоленская ФЛ",
    "Сокольники", "Университет 1,2", "Университет", "Чистые пруды", "Красносельская", "Павелецкая КЛ",
    "Арбатская АПЛ", "Лубянка 10", "Лубянка 11", "Лубянка", "Добрынинская", "Автозаводская", "Бауманская",
    "Октябрьская КЛ", "Площадь Революции 10,11", "Площадь Революции", "Курская АПЛ", "Охотный ряд 8,9", "Охотный ряд",
    "Боровицкая", "Ботанический сад", "Шаболовская", "Цветной бульвар", "Владыкино 1,2", "Владыкино",
    "Пушкинская", "Петровско-Разумовская", "Октябрьская КРЛ", "Динамо 1,2", "Динамо",
])


LEADERS_FULL = _dedupe_keep_order([
    "Хрипунов Александр Анатольевич",
    "Коваль Николай Сергеевич",
    "Киченко Александр Андреевич",
    "Васюк Леонид Николаевич",
    "Воронин Дмитрий Валерьевич",
    "Трофимов Олег Иванович",
    "Тимонин Антон Александрович",
    "Парфенов Максим Владимирович",
    "Матаев Антон Викторович",
    "Зиньковский Василий",
    "Тепляков Сергей",
    "Васильков Сергей Витальевич",
    "Максюткин И.О.",
])


_KNOWN_PHONE_SUPERVISORS = [
    ("Трофимов О. И.", "8(980)454-89-44"),
    ("Парфенов М. В.", "8(916)645-98-25"),
    ("Матаев А. В.", "8(915)310-20-80"),
]


def build_letter_supervisors():
    supervisors = list(_KNOWN_PHONE_SUPERVISORS)
    existing = {name for name, _ in supervisors}

    for full_name in LEADERS_FULL:
        short = fio_short(full_name)
        if short not in existing:
            supervisors.append((short, "не указан"))
            existing.add(short)

    return supervisors


LETTER_SUPERVISORS = build_letter_supervisors()
