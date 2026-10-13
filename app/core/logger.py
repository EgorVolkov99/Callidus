import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import sys

# Создаем папку для логов в хранилище
LOGS_DIR = Path(__file__).resolve().parent.parent.parent / "storage" / "logs"
LOGS_DIR.mkdir(parents=True, exist_ok=True)

LOG_FILE = LOGS_DIR / "callidus.log"

# Настраиваем форматтер (как будут выглядеть сообщения)
formatter = logging.Formatter(
    fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# Хэндлер для вывода в консоль
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)

# Хэндлер для записи в файл (максимум 5 МБ на файл, храним 3 последних)
file_handler = RotatingFileHandler(
    filename=LOG_FILE,
    mode='a',
    maxBytes=5 * 1024 * 1024,
    backupCount=3,
    encoding="utf-8"
)
file_handler.setFormatter(formatter)

# Настраиваем корневой логгер
logger = logging.getLogger("callidus")
logger.setLevel(logging.DEBUG)  # Ловим все сообщения уровня DEBUG и выше
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# Запрещаем дублирование логов от сторонних библиотек
logger.propagate = False