import os
import shutil
import re
import zipfile
import sys
import io
from datetime import datetime
from pathlib import Path

try:
    import py7zr
except ImportError:
    py7zr = None

# ==================== НАСТРОЙКА КОДИРОВКИ ====================

# Устанавливаем кодировку для консоли
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ==================== НАСТРОЙКИ ====================

# Базовый путь назначения
BASE_DEST_DIR = r"M:\Финансовый департамент\Treasury\Отчеты Брокера и Спецдепозитария"

RULES_XLSX_NAME = "Правила_разноса.xlsx"
RULES_XLSX_ALIASES = [
    "Правила разноса_актуальные.xlsx",
    "Правила_разноса.xlsx",
]

# Заполняются из Excel при запуске
SOURCES_CONFIG = []
FILE_RULES = []
ARCHIVE_RULES = []
ARCHIVE_FILE_RULES = []
MANAGER_MAPPING = {}
TKB_CODE_TO_PORTFOLIO = {}
SPUTNIK_VTB_CODES = {}
SPUTNIK_GPB_CODES = {}
SPUTNIK_RFB_CODES = {}

# Разрешенные расширения файлов
ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.xls', '.doc', '.docx', '.txt', '.xml', '.csv'}

# Что удаляем из названия при копировании
REMOVE_FROM_NAME = [
    "{Документооборот завершен}",
]

ZIP_METADATA_ENCODING = 'cp866'
ARCHIVE_EXTENSIONS = ('.zip', '.rar', '.7z')


def format_vtb_portfolio_folder(spurz_name, contract_code):
    """614BR6E1 -> SPURZ 1 (14BR6E_614BR6E0)"""
    if len(contract_code) < 3:
        return f"{spurz_name} ({contract_code})"
    inner = f"{contract_code[1:-1]}_{contract_code[:-1]}0"
    return f"{spurz_name} ({inner})"


def format_sputnik_portfolio_folder(code, broker=None):
    """Собирает имя папки портфеля СПУТНИК с номером договора."""
    code = str(code)
    broker = broker or ""

    if code in SPUTNIK_VTB_CODES or (code.startswith("614") and "ВТБ" in broker):
        spurz = SPUTNIK_VTB_CODES.get(code, "SPURZ")
        return format_vtb_portfolio_folder(spurz, code)

    if code in SPUTNIK_GPB_CODES or "ГПБ" in broker:
        spurz = SPUTNIK_GPB_CODES.get(code, "SPURZ")
        return f"{spurz} (Г000-Б-{code})"

    if code in SPUTNIK_RFB_CODES or "РФБ" in broker:
        spurz = SPUTNIK_RFB_CODES.get(code, "SPURZ")
        return f"{spurz} ({code})"

    if code.startswith("614") and len(code) >= 3:
        return format_vtb_portfolio_folder("SPURZ", code)

    return code


def get_rules_xlsx_path():
    """Ищет файл правил: сначала рядом со скриптом, потом в BASE_DEST_DIR."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = []
    for folder in (script_dir, BASE_DEST_DIR):
        for name in RULES_XLSX_ALIASES:
            candidates.append(os.path.join(folder, name))

    existing = [(p, os.path.getmtime(p)) for p in candidates if os.path.isfile(p)]
    if existing:
        existing.sort(key=lambda x: x[1], reverse=True)
        return existing[0][0]

    try:
        Path(BASE_DEST_DIR).mkdir(parents=True, exist_ok=True)
        return os.path.join(BASE_DEST_DIR, RULES_XLSX_NAME)
    except OSError:
        return os.path.join(script_dir, RULES_XLSX_NAME)


def load_rules():
    """Создаёт Excel с правилами при первом запуске и читает его."""
    global SOURCES_CONFIG, FILE_RULES, ARCHIVE_RULES, ARCHIVE_FILE_RULES
    global MANAGER_MAPPING, TKB_CODE_TO_PORTFOLIO
    global SPUTNIK_VTB_CODES, SPUTNIK_GPB_CODES, SPUTNIK_RFB_CODES

    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    from rules_excel import create_default_rules_xlsx, load_rules_from_xlsx

    xlsx_path = get_rules_xlsx_path()
    dest_dir = os.path.dirname(xlsx_path)
    try:
        Path(dest_dir).mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

    if not os.path.exists(xlsx_path):
        create_default_rules_xlsx(xlsx_path)

    data = load_rules_from_xlsx(xlsx_path)
    SOURCES_CONFIG = data["sources"]
    FILE_RULES = data["file_rules"]
    ARCHIVE_RULES = data["archive_rules"]
    ARCHIVE_FILE_RULES = data["archive_file_rules"]
    MANAGER_MAPPING = data["manager_mapping"]
    TKB_CODE_TO_PORTFOLIO = data["tkb_codes"]
    SPUTNIK_VTB_CODES = data["sputnik_vtb"]
    SPUTNIK_GPB_CODES = data["sputnik_gpb"]
    SPUTNIK_RFB_CODES = data["sputnik_rfb"]
    return xlsx_path


# ==================== ОСНОВНАЯ ЛОГИКА ====================


def get_manager_name(source_dir):
    """Извлекает имя управляющего из пути"""
    normalized_path = source_dir.replace('\\', '/').rstrip('/')
    raw_name = os.path.basename(normalized_path)
    return MANAGER_MAPPING.get(raw_name, raw_name)

def get_month_name(month_number):
    """Возвращает название месяца на русском"""
    months = {
        1: "Январь", 2: "Февраль", 3: "Март",
        4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь",
        10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(month_number, "")

def format_month_folder(date_obj):
    """Форматирует папку месяца: 06.Июнь 2026"""
    month_num = f"{date_obj.month:02d}"
    month_name = get_month_name(date_obj.month)
    year = date_obj.year
    return f"{month_num}.{month_name} {year}"

def clean_filename(filename):
    """Убирает префикс шлюза и плюсы из имени при копировании."""
    new_name = filename
    for remove_part in REMOVE_FROM_NAME:
        new_name = new_name.replace(remove_part, "")
    new_name = new_name.replace("+", "")
    return new_name


def parse_date_str(date_str, default_fmt):
    """Разбирает дату, учитывая 6 и 8 цифр."""
    candidates = [default_fmt]
    if len(date_str) == 8:
        candidates.extend(["%d%m%Y", "%Y%m%d"])
    elif len(date_str) == 6:
        candidates.extend(["%d%m%y", "%y%m%d"])
    seen = set()
    for fmt in candidates:
        if not fmt or fmt in seen:
            continue
        seen.add(fmt)
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def resolve_portfolio_folder(code, manager_name, broker=None):
    """Подставляет папку портфеля по коду из имени файла"""
    if not code:
        return None
    code = str(code)
    if manager_name == "ТКБ Инвестмент партнерс" or manager_name == "ТКБ":
        return TKB_CODE_TO_PORTFOLIO.get(code, code)
    if manager_name == "СПУТНИК":
        return format_sputnik_portfolio_folder(code, broker)
    return code

def extract_date_from_filename(filename, rule):
    """Извлекает дату из имени файла по правилу"""
    match = re.search(rule["date_regex"], filename, re.IGNORECASE)
    if not match:
        if rule.get("name_pattern") and filename.startswith(rule["name_pattern"]):
            name_part = filename[len(rule["name_pattern"]):]
            match = re.search(rule["date_regex"], name_part, re.IGNORECASE)
        if not match:
            return None, None, None, False

    groups = match.groupdict() if match.re.groupindex else {}
    date_str = groups.get("date")
    portfolio_code = groups.get("portfolio") or groups.get("code")
    contract_code = groups.get("contract")

    if not date_str:
        if rule.get("has_portfolio") and match.lastindex and match.lastindex >= 2:
            portfolio_code = match.group(1)
            date_str = match.group(2)
        elif rule.get("has_contract") and match.lastindex and match.lastindex >= 2:
            date_str = match.group(1)
            contract_code = match.group(2)
        elif match.lastindex:
            date_str = match.group(1)

    if not date_str:
        return None, None, None, False

    date_obj = parse_date_str(date_str, rule["date_format"])
    if not date_obj:
        return None, None, None, False
    return date_obj, portfolio_code, contract_code, True

def build_destination_path(date_obj, rule, manager_name, portfolio_code, original_filename, contract_code=None):
    """Строит путь назначения для обычного файла"""
    year = str(date_obj.year)
    month_folder = format_month_folder(date_obj)
    date_folder = date_obj.strftime("%d.%m.%Y")

    dest_path = rule["destination_template"]

    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    dest_path = dest_path.replace("*Дата*", date_folder)

    if portfolio_code:
        dest_path = dest_path.replace("*Портфель*", portfolio_code)
    if contract_code:
        dest_path = dest_path.replace("*Код договора*", str(contract_code))

    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    Path(full_path).mkdir(parents=True, exist_ok=True)

    clean_name = clean_filename(original_filename)
    if not clean_name or clean_name == os.path.splitext(original_filename)[1]:
        clean_name = original_filename

    return os.path.join(full_path, clean_name)

def normalize_filename(filename):
    """Сравнение имён: плюс, пробел, _, - и префикс шлюза не считаются отличием."""
    name = filename
    for part in REMOVE_FROM_NAME:
        name = name.replace(part, "")
    name_without_ext, ext = os.path.splitext(name)
    for char in ['+', ' ', '_', '-', '＋']:
        name_without_ext = name_without_ext.replace(char, '')
    return name_without_ext.lower() + ext.lower()

# Кэш содержимого папок назначения на один прогон (ускоряет проверку дублей)
_DEST_DIR_CACHE = {}


def clear_dest_cache():
    _DEST_DIR_CACHE.clear()


def check_file_exists(dest_path):
    """Проверяет, существует ли файл (с учётом + / пробелов / префикса шлюза)."""
    dest_dir = os.path.dirname(dest_path)
    target_filename = os.path.basename(dest_path)
    target_normalized = normalize_filename(target_filename)

    if dest_dir not in _DEST_DIR_CACHE:
        if not os.path.exists(dest_dir):
            _DEST_DIR_CACHE[dest_dir] = set()
        else:
            try:
                _DEST_DIR_CACHE[dest_dir] = {
                    normalize_filename(f)
                    for f in os.listdir(dest_dir)
                    if os.path.isfile(os.path.join(dest_dir, f))
                }
            except PermissionError:
                return False

    if target_normalized in _DEST_DIR_CACHE[dest_dir]:
        return True
    return False


def mark_file_written(dest_path):
    """Добавляет только что записанный файл в кэш дублей."""
    dest_dir = os.path.dirname(dest_path)
    name = normalize_filename(os.path.basename(dest_path))
    _DEST_DIR_CACHE.setdefault(dest_dir, set()).add(name)

# ==================== ОБРАБОТКА ОБЫЧНЫХ ФАЙЛОВ ====================

def process_regular_files(source_dir, manager_name, is_raiffaisen=False):
    """Обрабатывает обычные файлы"""
    if not os.path.exists(source_dir):
        return 0, 0

    files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
    if not files:
        return 0, 0

    processed = 0
    skipped = 0

    for filename in files:
        file_path = os.path.join(source_dir, filename)

        if filename.lower().endswith(ARCHIVE_EXTENSIONS):
            continue

        ext = os.path.splitext(filename)[1].lower()
        if ALLOWED_EXTENSIONS and ext and ext not in ALLOWED_EXTENSIONS:
            skipped += 1
            continue

        found_rule = None
        date_obj = None
        portfolio_code = None
        contract_code = None

        for rule in FILE_RULES:
            rule_manager = rule.get("manager") or ""
            if rule_manager and rule_manager != manager_name:
                continue
            date_obj, portfolio_code, contract_code, found = extract_date_from_filename(filename, rule)
            if found:
                found_rule = rule
                break

        if not found_rule or not date_obj:
            skipped += 1
            continue

        dest_path = build_destination_path(
            date_obj, found_rule, manager_name, portfolio_code, filename, contract_code
        )

        if check_file_exists(dest_path):
            skipped += 1
            continue

        try:
            shutil.copy2(file_path, dest_path)
            mark_file_written(dest_path)
            processed += 1
        except Exception:
            skipped += 1

    return processed, skipped

# ==================== ОБРАБОТКА АРХИВОВ ====================

def decode_zip_filename(name, flag_bits=0):
    """Восстанавливает имя файла из архива"""
    if flag_bits & 0x800:
        return name
    try:
        return name.encode('cp437').decode(ZIP_METADATA_ENCODING)
    except (UnicodeEncodeError, UnicodeDecodeError):
        return name

def open_zip(zip_path):
    """Открывает zip-архив с правильной кодировкой"""
    if sys.version_info >= (3, 11):
        try:
            return zipfile.ZipFile(zip_path, 'r', metadata_encoding=ZIP_METADATA_ENCODING), False
        except ValueError:
            pass
    return zipfile.ZipFile(zip_path, 'r'), True

def iter_archive_files(archive_path):
    """Возвращает пары (имя файла, bytes) из zip или 7z."""
    ext = os.path.splitext(archive_path)[1].lower()

    if ext == '.7z':
        if py7zr is None:
            raise RuntimeError("Для .7z нужен пакет py7zr: pip install py7zr")
        with py7zr.SevenZipFile(archive_path, mode='r') as archive:
            contents = archive.readall() or {}
            for name, bio in contents.items():
                if not name or name.endswith('/') or name.endswith('\\'):
                    continue
                base_name = os.path.basename(name.replace('\\', '/'))
                data = bio.read() if hasattr(bio, 'read') else bio
                yield base_name, data
        return

    if ext == '.rar':
        raise RuntimeError("Формат .rar пока не поддерживается, нужен .zip или .7z")

    zip_ref, need_decode = open_zip(archive_path)
    with zip_ref:
        for info in zip_ref.infolist():
            if info.is_dir():
                continue
            if need_decode:
                decoded_name = decode_zip_filename(info.filename, info.flag_bits)
            else:
                decoded_name = info.filename
            base_name = os.path.basename(decoded_name.replace('\\', '/'))
            yield base_name, zip_ref.read(info)

def extract_archive_info(filename):
    """Извлекает дату/портфель/брокера/период из имени архива."""
    for rule in ARCHIVE_RULES:
        pattern = rule.get("archive_pattern") or ""
        if pattern and not filename.startswith(pattern):
            continue

        match = re.search(rule["archive_regex"], filename, re.IGNORECASE)
        if not match:
            continue

        values = dict(match.groupdict()) if match.re.groupindex else {}
        if not values.get("date") and match.lastindex:
            values["date"] = match.group(1)
            if match.lastindex >= 2 and "portfolio" not in values:
                values["portfolio"] = match.group(2)

        date_str = values.get("date")
        if not date_str:
            continue

        date_obj = parse_date_str(date_str, rule.get("archive_date_format") or "%d%m%Y")
        if not date_obj:
            continue

        period = None
        if values.get("date") and values.get("date_end"):
            period = f"{values['date']}-{values['date_end']}"

        return {
            "date": date_obj,
            "portfolio": values.get("portfolio"),
            "broker": rule.get("broker"),
            "period": period,
            "kind": rule.get("kind"),
            "manager_override": rule.get("manager_override"),
            "extract_all": rule.get("extract_all", False),
            "destination_template": rule.get("destination_template"),
            "rule": rule,
        }

    return None

def extract_archive_file_info(filename):
    """Извлекает информацию из имени файла внутри архива"""
    for rule in ARCHIVE_FILE_RULES:
        match = re.search(rule["file_pattern"], filename, re.IGNORECASE)
        if not match:
            continue

        groups = match.groupdict() if match.re.groupindex else {}
        date_str = groups.get("date")
        broker_code = groups.get("code") or groups.get("contract")
        portfolio_code = groups.get("portfolio")

        if rule.get("no_portfolio"):
            broker_code = None
            portfolio_code = None

        if not date_str and match.lastindex:
            if rule.get("no_portfolio"):
                date_str = match.group(1)
            elif match.lastindex >= 2:
                broker_code = broker_code or match.group(1)
                date_str = match.group(2)
            else:
                date_str = match.group(1)

        if not date_str:
            continue

        date_obj = parse_date_str(date_str, rule.get("file_date_format") or "%Y%m%d")
        if not date_obj:
            continue
        return date_obj, portfolio_code, broker_code, rule
    return None, None, None, None

def build_archive_destination_path(
    date_obj, rule, manager_name, portfolio_code, broker_code, filename,
    broker=None, period=None, contract_code=None
):
    """Строит путь назначения для файла из архива"""
    year = str(date_obj.year)
    month_folder = format_month_folder(date_obj)
    date_folder = date_obj.strftime("%d.%m.%Y")

    dest_path = rule["destination_template"]
    portfolio_folder = portfolio_code if portfolio_code else broker_code

    dest_path = dest_path.replace("*ГГГГ*", year)
    dest_path = dest_path.replace("*Месяц*", month_folder)
    dest_path = dest_path.replace("*Управляющий*", manager_name)
    dest_path = dest_path.replace("*Дата*", date_folder)

    if portfolio_folder:
        dest_path = dest_path.replace("*Портфель*", str(portfolio_folder))
    if broker:
        dest_path = dest_path.replace("*Брокер*", broker)
    if period:
        dest_path = dest_path.replace("*Период*", period)
    if contract_code:
        dest_path = dest_path.replace("*Код договора*", str(contract_code))

    # Если плейсхолдеры остались незаполненными — убираем их сегменты грубо
    dest_path = dest_path.replace("*Брокер*\\", "").replace("*Брокер*", "")
    dest_path = dest_path.replace("*Портфель*\\", "").replace("*Портфель*", "")
    dest_path = dest_path.replace("*Период*\\", "").replace("*Период*", "")
    dest_path = dest_path.replace("*Код договора*\\", "").replace("*Код договора*", "")

    full_path = os.path.join(BASE_DEST_DIR, dest_path)
    Path(full_path).mkdir(parents=True, exist_ok=True)

    return os.path.join(full_path, filename)

def process_archives(source_dir, manager_name):
    """Обрабатывает zip/7z архивы"""
    if not os.path.exists(source_dir):
        return 0, 0

    files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
    archives = [f for f in files if f.lower().endswith(ARCHIVE_EXTENSIONS)]
    if not archives:
        return 0, 0

    processed = 0
    skipped = 0

    for archive_name in archives:
        archive_path = os.path.join(source_dir, archive_name)
        archive_info = extract_archive_info(archive_name)
        if not archive_info:
            skipped += 1
            continue

        archive_date = archive_info["date"]
        portfolio_code = archive_info.get("portfolio")
        broker = archive_info.get("broker")
        period = archive_info.get("period")
        effective_manager = archive_info.get("manager_override") or manager_name

        try:
            members = list(iter_archive_files(archive_path))

            for base_name, file_data in members:
                ext = os.path.splitext(base_name)[1].lower()
                if ALLOWED_EXTENSIONS and ext not in ALLOWED_EXTENSIONS:
                    skipped += 1
                    continue

                if archive_info.get("extract_all") and archive_info.get("destination_template"):
                    fake_rule = {
                        "destination_template": archive_info["destination_template"],
                        "identifier": "tkb_napf_extract_all",
                    }
                    dest_path = build_archive_destination_path(
                        archive_date, fake_rule, effective_manager,
                        portfolio_code, None, base_name,
                        broker=broker, period=period
                    )
                    if check_file_exists(dest_path):
                        skipped += 1
                        continue
                    try:
                        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                        with open(dest_path, 'wb') as f:
                            f.write(file_data)
                        mark_file_written(dest_path)
                        processed += 1
                    except Exception:
                        skipped += 1
                    continue

                file_date, file_portfolio, broker_code, rule = extract_archive_file_info(base_name)

                if not file_date:
                    skipped += 1
                    continue

                final_portfolio = portfolio_code if portfolio_code else file_portfolio
                contract_code = None

                if rule.get("use_code_map") and broker_code:
                    mapped = resolve_portfolio_folder(broker_code, effective_manager, broker)
                    final_portfolio = mapped
                    contract_code = broker_code
                elif rule["identifier"] in ["trades", "brok_rpt"]:
                    final_portfolio = broker_code

                dest_rule = rule
                if broker and (rule.get("needs_broker") or rule.get("use_code_map")):
                    dest_rule = dict(rule)
                    dest_rule["destination_template"] = (
                        r"2 Отчеты брокера\*ГГГГ*\*Месяц*\*Управляющий*\*Брокер*\*Портфель*"
                    )

                if not rule.get("no_portfolio") and not final_portfolio:
                    skipped += 1
                    continue

                dest_path = build_archive_destination_path(
                    file_date, dest_rule, effective_manager,
                    final_portfolio, broker_code, base_name,
                    broker=broker, period=period, contract_code=contract_code
                )

                if check_file_exists(dest_path):
                    skipped += 1
                    continue

                try:
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    with open(dest_path, 'wb') as f:
                        f.write(file_data)
                    mark_file_written(dest_path)
                    processed += 1
                except Exception:
                    skipped += 1

        except Exception:
            skipped += 1

    return processed, skipped

# ==================== ОСНОВНАЯ ФУНКЦИЯ ====================

def process_all_sources():
    """Обрабатывает все источники"""
    try:
        rules_path = load_rules()
    except Exception as e:
        print(f"Ошибка загрузки правил: {e}")
        return

    if not SOURCES_CONFIG:
        print("В Excel нет источников (лист «Источники»).")
        return

    clear_dest_cache()
    total_processed = 0
    total_skipped = 0

    for source_config in SOURCES_CONFIG:
        source_dir = source_config["source_dir"]
        manager_name = source_config.get("manager") or get_manager_name(source_dir)
        source_type = source_config.get("type", "files")

        if source_type == "files":
            processed, skipped = process_regular_files(source_dir, manager_name)
        elif source_type == "archives":
            processed, skipped = process_archives(source_dir, manager_name)
        elif source_type == "mixed":
            processed1, skipped1 = process_regular_files(source_dir, manager_name)
            processed2, skipped2 = process_archives(source_dir, manager_name)
            processed = processed1 + processed2
            skipped = skipped1 + skipped2
        else:
            continue

        total_processed += processed
        total_skipped += skipped

    print("Готово")
    print(f"Скопировано: {total_processed}")
    print(f"Пропущено: {total_skipped}")
    print(f"Назначение: {BASE_DEST_DIR}")
    print(f"Правила: {rules_path}")

# ==================== ЗАПУСК ====================

if __name__ == "__main__":
    process_all_sources()
