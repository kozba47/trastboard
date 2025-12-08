# config.py

from pathlib import Path

# Путь к Excel-файлу с данными для дашборда.
# Для начала кладём файл рядом с app.py под именем "dashboard_data.xlsx".
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "data" / "dashboard_data.xlsx"

# Здесь можно позже добавить другие настройки:
# - дефолтный порядок блоков
# - маппинг "лист -> тип блока"
# - параметры кэша чтения Excel и т.д.
