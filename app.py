# -*- coding: utf-8 -*-
"""
==========================================================================
  ГЕНЕРАТОР СОР / СОЧ ПО ИСТОРИИ  (Республика Казахстан)
==========================================================================

Необходимые библиотеки (установка одной командой):

    pip install streamlit google-genai pillow python-docx requests

Запуск:

    streamlit run app.py

Секреты / переменные окружения (st.secrets или os.environ):

    APP_PASSWORD       - daniyar
    GEMINI_API_KEY     - AQ.Ab8RN6K8NjKcVd2Tv1m--86AZQ6a7UwRYFQ0-lTxr4-iMqKm_A
    TELEGRAM_BOT_TOKEN - токен Telegram-бота (необязательно)
    TELEGRAM_CHAT_ID   - ID чата для уведомлений (необязательно)
==========================================================================
"""

from __future__ import annotations

import io
import os
import re
import tempfile
import time
from datetime import datetime
from pathlib import Path

import requests
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor
from google import genai
from google.genai import types
from PIL import Image

# --------------------------------------------------------------------------
# 0. БАЗОВАЯ НАСТРОЙКА СТРАНИЦЫ (wide-layout)
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Генератор СОР и СОЧ",
    page_icon="📜",
    layout="wide",
    initial_sidebar_state="expanded",
)

MODEL_NAME = "gemini-2.5-flash"
DEFAULT_PASSWORD = "admin123"
IMAGE_MARKER = "[ИЗОБРАЖЕНИЕ]"  # метка, куда вставляется картинка в Word

# Папка с учебниками рядом с app.py. Структура:
#   textbooks/История Казахстана/7/учебник.pdf
#   textbooks/Всемирная история/7/учебник.pdf
TEXTBOOK_DIR = Path(__file__).parent / "textbooks"
FALLBACK_MODELS = ["gemini-2.5-flash-lite"]  # запасная модель при перегрузке основной
MAX_TEXTBOOKS = 3  # сколько учебников максимум отправлять в одном запросе


# --------------------------------------------------------------------------
# 1. ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ДЛЯ СЕКРЕТОВ
# --------------------------------------------------------------------------
def get_secret(name: str, default: str = "") -> str:
    """Сначала ищем в st.secrets, затем в переменных окружения, затем default."""
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        # secrets.toml может отсутствовать - это нормально
        pass
    return os.environ.get(name, default)


# --------------------------------------------------------------------------
# 2. КАСТОМНЫЙ CSS (минималистичный светлый стиль, как у ChatGPT / Claude)
# --------------------------------------------------------------------------
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Lora:wght@600;700&display=swap');

    /* ---------- Палитра (тёплая бумажная + индиго) ---------- */
    :root {
        color-scheme: light only;
        --bg: #FBF8F3;
        --panel: #F2EDE4;
        --card: #FFFFFF;
        --ink: #1E2230;
        --muted: #6A6F7D;
        --line: #DDD6C8;
        --accent: #3B4CCA;
        --accent-dark: #2D3AA3;
        --accent-soft: #E8EAFB;
    }

    /* Принудительно светлая схема, даже если в системе включена тёмная */
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        background: var(--bg) !important;
        color: var(--ink) !important;
        font-family: 'Inter', -apple-system, "Segoe UI", sans-serif !important;
    }
    header[data-testid="stHeader"], [data-testid="stToolbar"] {
        background: var(--bg) !important;
    }
    header[data-testid="stHeader"] * { color: var(--ink) !important; }
    #MainMenu, footer { visibility: hidden; }

    .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1000px; }

    /* ---------- Весь текст и подписи ---------- */
    .stApp p, .stApp span, .stApp li, .stApp label, .stApp div,
    [data-testid="stWidgetLabel"] p, [data-testid="stMarkdownContainer"] p {
        color: var(--ink);
    }
    h1, h2, h3, h4 {
        font-family: 'Lora', Georgia, serif !important;
        color: var(--ink) !important;
        letter-spacing: -0.01em;
    }
    h1 { font-size: 2rem !important; }
    [data-testid="stWidgetLabel"] p {
        font-weight: 600; font-size: .88rem; color: var(--ink) !important;
    }

    /* ---------- Боковая панель ---------- */
    section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div {
        background: var(--panel) !important;
        border-right: 1px solid var(--line);
    }
    section[data-testid="stSidebar"] * { color: var(--ink); }

    /* ---------- Поля ввода: тёмный текст на белом ---------- */
    .stTextInput input, .stTextArea textarea, .stNumberInput input {
        background: #FFFFFF !important;
        color: var(--ink) !important;
        -webkit-text-fill-color: var(--ink) !important;
        caret-color: var(--accent);
        border: 1.5px solid var(--line) !important;
        border-radius: 12px !important;
        padding: .7rem .9rem !important;
    }
    .stTextInput input::placeholder, .stTextArea textarea::placeholder {
        color: #9A9EAB !important;
        -webkit-text-fill-color: #9A9EAB !important;
        opacity: 1;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: var(--accent) !important;
        box-shadow: 0 0 0 3px var(--accent-soft) !important;
    }
    [data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="base-input"] {
        background: #FFFFFF !important;
        border-radius: 12px !important;
    }

    /* ---------- Выпадающие списки ---------- */
    div[data-baseweb="select"] > div {
        background: #FFFFFF !important;
        color: var(--ink) !important;
        border: 1.5px solid var(--line) !important;
        border-radius: 12px !important;
    }
    div[data-baseweb="select"] * { color: var(--ink) !important; }
    div[data-baseweb="select"] svg { fill: var(--muted) !important; }
    div[data-baseweb="popover"], div[data-baseweb="popover"] ul,
    ul[role="listbox"] {
        background: #FFFFFF !important;
    }
    ul[role="listbox"] li, ul[role="listbox"] li * { color: var(--ink) !important; }
    ul[role="listbox"] li:hover, li[aria-selected="true"] {
        background: var(--accent-soft) !important;
    }

    /* ---------- Слайдер ---------- */
    [data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
        background: var(--accent) !important;
    }
    [data-testid="stSlider"] [data-testid="stThumbValue"],
    [data-testid="stSlider"] [data-testid="stTickBarMin"],
    [data-testid="stSlider"] [data-testid="stTickBarMax"] {
        color: var(--accent) !important;
    }

    /* ---------- Загрузка файлов ---------- */
    [data-testid="stFileUploaderDropzone"] {
        background: #FFFFFF !important;
        border: 1.5px dashed var(--line) !important;
        border-radius: 12px !important;
    }
    [data-testid="stFileUploaderDropzone"] * { color: var(--ink) !important; }
    [data-testid="stFileUploaderDropzone"] button {
        background: var(--accent-soft) !important;
        color: var(--accent-dark) !important;
        border: none !important;
    }

    /* ---------- Чекбокс ---------- */
    [data-testid="stCheckbox"] label span { color: var(--ink) !important; }

    /* ---------- Кнопки ---------- */
    .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
        background: var(--accent) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 12px !important;
        padding: .7rem 1.3rem !important;
        font-weight: 600 !important;
        width: 100%;
        transition: all .15s ease;
    }
    .stButton > button *, .stDownloadButton > button *, .stFormSubmitButton > button * {
        color: #FFFFFF !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover,
    .stFormSubmitButton > button:hover {
        background: var(--accent-dark) !important;
        box-shadow: 0 6px 16px rgba(59,76,202,.28);
        transform: translateY(-1px);
    }
    /* Кнопка выхода - спокойная */
    section[data-testid="stSidebar"] .stButton > button {
        background: #FFFFFF !important;
        border: 1.5px solid var(--line) !important;
    }
    section[data-testid="stSidebar"] .stButton > button * { color: var(--ink) !important; }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: #F6F2EA !important; box-shadow: none; transform: none;
    }

    /* ---------- Верхний баннер ---------- */
    .hero {
        background: linear-gradient(135deg, #3B4CCA 0%, #6A5BD6 100%);
        border-radius: 18px;
        padding: 1.6rem 1.8rem;
        margin-bottom: 1.4rem;
        box-shadow: 0 10px 28px rgba(59,76,202,.22);
    }
    .hero h1, .hero p { color: #FFFFFF !important; margin: 0; }
    .hero h1 { font-size: 1.8rem !important; }
    .hero p { opacity: .9; margin-top: .35rem; font-size: .95rem; }

    /* ---------- Карточка результата ---------- */
    .result-card {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 16px;
        padding: 1.6rem 1.8rem;
        box-shadow: 0 2px 10px rgba(30,34,48,.05);
        margin-top: 1rem;
    }
    [data-testid="stMarkdownContainer"] table { border-collapse: collapse; width: 100%; }
    [data-testid="stMarkdownContainer"] th, [data-testid="stMarkdownContainer"] td {
        border: 1px solid var(--line); padding: .45rem .6rem; color: var(--ink);
    }
    [data-testid="stMarkdownContainer"] th { background: var(--panel); }

    /* ---------- Экран входа ---------- */
    .login-title { text-align: center; margin-top: 3.5rem; }
    .login-sub { text-align: center; color: var(--muted) !important; margin-bottom: 1.5rem; }
    .small-muted { color: var(--muted) !important; font-size: .88rem; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# 3. АВТОРИЗАЦИЯ
# --------------------------------------------------------------------------
def check_password() -> bool:
    """Показывает форму входа. Возвращает True, если пользователь вошёл."""
    if st.session_state.get("authenticated"):
        return True

    correct_password = get_secret("APP_PASSWORD", DEFAULT_PASSWORD)

    st.markdown("<h1 class='login-title'>📜 Генератор СОР и СОЧ</h1>", unsafe_allow_html=True)
    st.markdown(
        "<p class='login-sub'>История Казахстана и Всемирная история · критериальное оценивание</p>",
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1, 1.2, 1])
    with center:
        with st.form("login_form"):
            password = st.text_input("Пароль", type="password", placeholder="Введите пароль")
            submitted = st.form_submit_button("Войти")
        if submitted:
            if password == correct_password:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Неверный пароль. Попробуйте ещё раз.")
    return False


# --------------------------------------------------------------------------
# 4. СИСТЕМНЫЙ ПРОМПТ И ПОСТРОЕНИЕ ЗАПРОСА
# --------------------------------------------------------------------------
SYSTEM_PROMPT = """Ты - опытный методист по критериальному оцениванию в школах Республики Казахстан.
Ты отлично знаешь обновлённую программу по истории Казахстана и всемирной истории, 
структуру СОР (суммативное оценивание за раздел) и СОЧ (суммативное оценивание за четверть),
а также таксономию мыслительных навыков и требования к составлению критериев и дескрипторов.

ТВОЯ ЗАДАЧА: составить готовую к печати работу для учителя.

ОБЯЗАТЕЛЬНЫЕ ТРЕБОВАНИЯ:
1. Задания должны быть РАЗНОГО УРОВНЯ мыслительных навыков и идти по нарастающей сложности:
   - Знание и понимание (факты, даты, персоналии, термины);
   - Применение (использование знаний в новой ситуации, работа с источником/картой);
   - Анализ (причины, следствия, сравнение, работа с историческим источником);
   - Синтез и оценка (аргументация, собственная позиция, вывод, эссе-рассуждение).
   Около половины заданий должны проверять навыки высокого порядка.
2. К КАЖДОМУ заданию обязательно указывай:
   - уровень мыслительных навыков;
   - критерий оценивания (кратко);
   - дескрипторы (проверяемые действия обучающегося, формулировка «Обучающийся ...») 
     с баллами за каждый дескриптор;
   - максимальный балл за задание.
3. В КОНЦЕ работы обязательно дай раздел «Схема выставления баллов (модель ответов)»: 
   в виде таблицы Markdown с колонками «№ задания | Ответ / возможные варианты ответа | Балл | Дополнительная информация», 
   и итоговый общий балл.
4. Содержание должно строго соответствовать теме и классу, быть исторически точным, 
   без выдуманных фактов, дат и цитат. Если не уверен в факте - не используй его.
5. Язык заданий - русский. Формулировки чёткие, соответствующие возрасту обучающихся.
6. Для СОР: небольшая работа (примерно 20 минут, 10-20 баллов), включает задания разных типов 
   (тест с выбором ответа, соответствие, краткий ответ, развёрнутый ответ).
   Для СОЧ: более объёмная работа (примерно 40 минут, 25-40 баллов), охватывает материал всей четверти, 
   включает разделы заданий с нарастающей сложностью.
7. Если несколько вариантов - чётко разделяй их заголовками «Вариант 1», «Вариант 2» 
   и давай схему баллов для каждого.

ФОРМАТ ВЫВОДА (строго Markdown):
- Заглавие работы: «# ...» (в нём тип работы, предмет, класс, тема/четверть).
- Строка для данных ученика: «Ф.И.О. обучающегося: ________  Класс: ____  Дата: ____».
- Разделы через «## ...», задания как «### Задание N (X баллов)».
- Дескрипторы - маркированным списком «- ...» с баллом в конце, например «(1 балл)».
- Жирный текст через **...**. Таблицы - обычные Markdown-таблицы с «|».
- Не используй эмодзи. Не добавляй вступлений и пояснений вне самой работы.
"""


def build_user_prompt(
    work_type: str,
    subject: str,
    grade: int,
    topic: str,
    quarter: int,
    tasks_count: int,
    variants: int,
    extra: str,
    has_image: bool,
    has_books: bool = False,
) -> str:
    """Формирует текст запроса пользователя к модели."""
    work_full = (
        "СОР (Суммативное оценивание за раздел)"
        if work_type == "СОР"
        else "СОЧ (Суммативное оценивание за четверть)"
    )
    lines = [
        f"Составь {work_full}.",
        f"Предмет: {subject}.",
        f"Класс: {grade}.",
    ]
    if work_type == "СОР":
        lines.append(f"Раздел / тема: {topic}.")
    else:
        lines.append(f"Четверть: {quarter}. Охватываемые темы: {topic}.")
    lines.append(f"Количество заданий в одном варианте: примерно {tasks_count}.")
    lines.append(f"Количество вариантов: {variants}.")

    if has_image:
        lines.append(
            "К запросу ПРИКРЕПЛЕНА историческая иллюстрация (карта, портрет, документ или рисунок). "
            "Внимательно изучи её и ОБЯЗАТЕЛЬНО встрой в одно из заданий (например, анализ источника, "
            "работа с картой, определение личности/события по изображению). "
            f"В тексте задания на том месте, где должна находиться картинка, поставь метку {IMAGE_MARKER} "
            "отдельной строкой. Не описывай картинку слишком подробно, чтобы не выдать ответ."
        )
    if has_books:
        lines.append(
            "К запросу ПРИКРЕПЛЕНЫ школьные учебники (PDF). Опирайся на их содержание: "
            "термины, факты, даты и имена бери из учебников, соответствующих теме и классу. "
            "Не выходи за рамки материала учебников и не выдумывай факты, которых в них нет. "
            "Если нужной темы в приложенных учебниках нет, скажи об этом одной строкой в начале работы."
        )
    if extra.strip():
        lines.append(f"Дополнительные пожелания учителя: {extra.strip()}")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# 5. ГЕНЕРАЦИЯ ЧЕРЕЗ GEMINI (google-genai)
# --------------------------------------------------------------------------
def list_textbooks(subject: str, grade: int) -> list:
    """Возвращает PDF-учебники из textbooks/<предмет>/<класс>/."""
    folder = TEXTBOOK_DIR / subject / str(grade)
    if not folder.is_dir():
        return []
    return sorted(folder.glob("*.pdf"))


def save_uploaded_textbooks(files, subject: str, grade: int) -> int:
    """Сохраняет загруженные в приложении PDF в библиотеку (они останутся навсегда)."""
    folder = TEXTBOOK_DIR / subject / str(grade)
    folder.mkdir(parents=True, exist_ok=True)
    saved = 0
    for f in files:
        target = folder / Path(f.name).name
        target.write_bytes(f.getbuffer())
        saved += 1
    return saved


@st.cache_data(ttl=60 * 60 * 20, show_spinner=False)
def upload_textbook(path_str: str, mtime: float, api_key: str):
    """Загружает PDF в Gemini Files API один раз и кэширует ссылку на ~20 часов
    (Gemini хранит файлы 48 часов). mtime нужен, чтобы кэш обновлялся при замене файла."""
    client = genai.Client(api_key=api_key)
    uploaded = client.files.upload(
        file=path_str,
        config=types.UploadFileConfig(mime_type="application/pdf"),
    )
    # Ждём, пока файл обработается
    for _ in range(60):
        info = client.files.get(name=uploaded.name)
        state = str(getattr(info, "state", ""))
        if "ACTIVE" in state:
            return info.uri, info.mime_type or "application/pdf"
        if "FAILED" in state:
            raise RuntimeError(f"Не удалось обработать учебник: {Path(path_str).name}")
        time.sleep(2)
    raise RuntimeError(f"Учебник обрабатывается слишком долго: {Path(path_str).name}")


def generate_work(
    user_prompt: str, image: Image.Image | None, book_paths: list | None = None
) -> str:
    """Отправляет запрос в Gemini и возвращает текст работы."""
    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Не найден GEMINI_API_KEY. Добавьте ключ в st.secrets или в переменные окружения."
        )

    client = genai.Client(api_key=api_key)

    # Контент: текст + (при наличии) картинка PIL
    contents: list = []
    # Учебники (PDF) идут первыми, затем текст запроса и картинка
    for bp in (book_paths or [])[:MAX_TEXTBOOKS]:
        uri, mime = upload_textbook(str(bp), Path(bp).stat().st_mtime, api_key)
        contents.append(types.Part.from_uri(file_uri=uri, mime_type=mime))
    contents.append(user_prompt)
    if image is not None:
        contents.append(image)

    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        temperature=0.7,
    )

    # Серверы Google иногда перегружены (ошибки 503/429). Делаем несколько попыток
    # с паузами, а если основная модель не отвечает - пробуем запасную.
    response = None
    last_error = None
    st.session_state["used_model"] = MODEL_NAME
    for model in [MODEL_NAME] + FALLBACK_MODELS:
        for attempt in range(4):
            try:
                response = client.models.generate_content(
                    model=model, contents=contents, config=config
                )
                st.session_state["used_model"] = model
                break
            except Exception as e:  # noqa: BLE001
                last_error = e
                msg = str(e)
                temporary = any(
                    k in msg for k in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "500", "overloaded", "high demand")
                )
                if not temporary:
                    raise  # настоящая ошибка (неверный ключ и т.п.) - сразу показываем
                time.sleep(3 * (attempt + 1))  # 3, 6, 9, 12 секунд
        if response is not None:
            break
    if response is None:
        raise RuntimeError(
            "Серверы Google сейчас перегружены. Подождите 1-2 минуты и нажмите кнопку ещё раз. "
            f"(Детали: {last_error})"
        )
    text = (response.text or "").strip()
    if not text:
        raise RuntimeError("Модель вернула пустой ответ. Попробуйте изменить тему и повторить.")
    return text


# --------------------------------------------------------------------------
# 6. УВЕДОМЛЕНИЕ В TELEGRAM
# --------------------------------------------------------------------------
def send_telegram_notification(subject: str, grade: int, work_type: str, topic: str) -> None:
    """Отправляет сообщение в Telegram. Ошибки не прерывают работу приложения."""
    token = get_secret("TELEGRAM_BOT_TOKEN")
    chat_id = get_secret("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return  # Telegram не настроен - молча пропускаем

    text = (
        "✅ Сгенерирована новая работа\n"
        f"Тип: {work_type}\n"
        f"Предмет: {subject}\n"
        f"Класс: {grade}\n"
        f"Тема: {topic}\n"
        f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M')}"
    )
    try:
        requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data={"chat_id": chat_id, "text": text},
            timeout=10,
        )
    except Exception:
        # Уведомление - вторичная функция, не мешаем основной работе
        pass


# --------------------------------------------------------------------------
# 7. ЭКСПОРТ В WORD (.docx)
# --------------------------------------------------------------------------
def add_formatted_runs(paragraph, text: str) -> None:
    """Разбирает **жирный** текст и добавляет фрагменты в абзац."""
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part.replace("*", ""))


def prepare_image_stream(image: Image.Image) -> io.BytesIO:
    """Конвертирует PIL-изображение в PNG-поток для вставки в Word."""
    img = image.convert("RGB")
    stream = io.BytesIO()
    img.save(stream, format="PNG")
    stream.seek(0)
    return stream


def insert_picture(doc: Document, image: Image.Image) -> None:
    """Вставляет картинку по центру."""
    doc.add_picture(prepare_image_stream(image), width=Cm(12))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER


def build_docx(text: str, image: Image.Image | None, header_line: str) -> bytes:
    """Преобразует Markdown-подобный текст модели в аккуратный Word-документ."""
    doc = Document()

    # Поля страницы
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)

    # Базовый стиль
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    image_inserted = False
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.strip()

        # Пустая строка
        if not line:
            i += 1
            continue

        # Метка изображения
        if IMAGE_MARKER in line:
            if image is not None and not image_inserted:
                insert_picture(doc, image)
                image_inserted = True
            i += 1
            continue

        # Таблица Markdown (подряд идущие строки с "|")
        if line.startswith("|"):
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row_line = lines[i].strip()
                # пропускаем разделитель |---|---|
                if not re.fullmatch(r"\|[\s:\-|]+\|?", row_line):
                    cells = [c.strip() for c in row_line.strip("|").split("|")]
                    table_rows.append(cells)
                i += 1
            if table_rows:
                cols = max(len(r) for r in table_rows)
                table = doc.add_table(rows=0, cols=cols)
                table.style = "Table Grid"
                for r_idx, row in enumerate(table_rows):
                    cells = table.add_row().cells
                    for c_idx in range(cols):
                        cell_text = row[c_idx] if c_idx < len(row) else ""
                        para = cells[c_idx].paragraphs[0]
                        add_formatted_runs(para, cell_text.replace("<br>", "\n"))
                        if r_idx == 0:  # шапка - жирным
                            for run in para.runs:
                                run.bold = True
                doc.add_paragraph()
            continue

        # Заголовки
        heading_match = re.match(r"^(#{1,4})\s+(.*)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            title = heading_match.group(2).replace("**", "").strip()
            if level == 1:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(title)
                run.bold = True
                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
            else:
                heading = doc.add_heading(title, level=min(level, 3))
                for run in heading.runs:
                    run.font.name = "Times New Roman"
                    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
            i += 1
            continue

        # Маркированный список
        bullet_match = re.match(r"^[-*•]\s+(.*)$", line)
        if bullet_match:
            p = doc.add_paragraph(style="List Bullet")
            add_formatted_runs(p, bullet_match.group(1))
            i += 1
            continue

        # Горизонтальная линия
        if re.fullmatch(r"[-_*]{3,}", line):
            i += 1
            continue

        # Обычный абзац
        p = doc.add_paragraph()
        add_formatted_runs(p, line)
        i += 1

    # Если модель не поставила метку, а картинка есть - добавим в конец, чтобы не потерять
    if image is not None and not image_inserted:
        doc.add_paragraph()
        cap = doc.add_paragraph()
        cap.add_run("Иллюстрация к заданию:").bold = True
        insert_picture(doc, image)

    # Колонтитул с краткой информацией
    footer_p = doc.sections[0].footer.paragraphs[0]
    footer_p.text = header_line
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer_p.runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# 8. ОСНОВНОЙ ИНТЕРФЕЙС
# --------------------------------------------------------------------------
def main_app() -> None:
    # ---------------- Боковая панель ----------------
    with st.sidebar:
        st.markdown("### 📜 Параметры работы")

        work_type = st.selectbox(
            "Тип работы",
            options=["СОР", "СОЧ"],
            format_func=lambda x: "СОР (Раздел)" if x == "СОР" else "СОЧ (Четверть)",
        )
        subject = st.selectbox("Предмет", ["История Казахстана", "Всемирная история"])
        grade = st.selectbox("Класс", list(range(5, 12)), index=2)

        quarter = 1
        if work_type == "СОЧ":
            quarter = st.selectbox("Четверть", [1, 2, 3, 4])

        tasks_count = st.slider(
            "Количество заданий",
            min_value=4,
            max_value=12,
            value=6 if work_type == "СОР" else 9,
        )
        variants = st.selectbox("Количество вариантов", [1, 2], index=0)

        st.markdown("---")
        uploaded = st.file_uploader(
            "Изображение к заданию (карта, портрет, иллюстрация)",
            type=["png", "jpg", "jpeg"],
            help="Картинка будет встроена в одно из заданий и в Word-файл.",
        )
        if uploaded is not None:
            st.image(uploaded, caption="Прикреплено", use_container_width=True)

        st.markdown("---")
        st.markdown("**📚 Учебники**")
        book_paths = list_textbooks(subject, grade)
        use_books = st.checkbox(
            "Опираться на учебники",
            value=bool(book_paths),
            help="Задания составляются по содержанию учебников из библиотеки.",
        )
        if book_paths:
            st.caption(f"В библиотеке ({subject}, {grade} класс): " + ", ".join(b.name for b in book_paths))
        else:
            st.caption(f"Для «{subject}», {grade} класс учебников пока нет.")
        new_books = st.file_uploader(
            "Добавить учебник (PDF)", type=["pdf"], accept_multiple_files=True, key="books_up"
        )
        if new_books and st.button("Сохранить в библиотеку"):
            n = save_uploaded_textbooks(new_books, subject, grade)
            st.success(f"Сохранено: {n}. Обновите страницу.")

        st.markdown("---")
        if st.button("Выйти из аккаунта"):
            st.session_state.clear()
            st.rerun()

    # ---------------- Центральная часть ----------------
    st.markdown(
        "<div class='hero'><h1>Генератор СОР и СОЧ</h1>"
        "<p>Критериальное оценивание · дескрипторы · схема баллов · экспорт в Word</p></div>",
        unsafe_allow_html=True,
    )

    topic_label = "Тема / раздел" if work_type == "СОР" else "Темы четверти (через запятую)"
    topic_placeholder = (
        "Например: Образование Казахского ханства"
        if work_type == "СОР"
        else "Например: Казахское ханство, Золотая Орда, Жузы"
    )
    topic = st.text_input(topic_label, placeholder=topic_placeholder)
    extra = st.text_area(
        "Дополнительные пожелания (необязательно)",
        placeholder="Например: добавить задание на работу с датами, упростить формулировки…",
        height=90,
    )

    generate_clicked = st.button("✨ Сгенерировать работу")

    if generate_clicked:
        if not topic.strip():
            st.warning("Пожалуйста, укажите тему.")
        else:
            image = None
            if uploaded is not None:
                try:
                    image = Image.open(uploaded)
                    image.load()
                except Exception:
                    st.error("Не удалось прочитать изображение. Загрузите корректный PNG или JPG.")
                    return

            prompt = build_user_prompt(
                work_type, subject, grade, topic, quarter, tasks_count, variants, extra,
                image is not None, bool(use_books and book_paths),
            )

            try:
                with st.spinner("Генерирую работу… с учебниками это может занять 1–3 минуты"):
                    result_text = generate_work(
                        prompt, image, book_paths if use_books else None
                    )

                # Формируем Word сразу после успешной генерации
                footer_line = f"{work_type} · {subject} · {grade} класс · {topic}"
                docx_bytes = build_docx(result_text, image, footer_line)

                # Сохраняем в session_state, чтобы результат не пропадал при перерисовке
                st.session_state["result"] = {
                    "text": result_text,
                    "docx": docx_bytes,
                    "file_name": f"{work_type}_{subject.replace(' ', '_')}_{grade}_класс.docx",
                    "image": image,
                }

                # Мгновенное уведомление в Telegram
                send_telegram_notification(subject, grade, work_type, topic)

            except Exception as e:
                st.error(f"Ошибка генерации: {e}")

    # ---------------- Показ результата ----------------
    result = st.session_state.get("result")
    if result:
        st.success("Работа готова. Просмотрите её ниже и скачайте в формате Word.")
        if st.session_state.get("used_model") not in (None, MODEL_NAME):
            st.info("Основная модель была перегружена, работа создана запасной моделью. Проверьте её особенно внимательно.")

        st.download_button(
            label="📥 Скачать документ в формате Word (.docx)",
            data=result["docx"],
            file_name=result["file_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.markdown("<div class='result-card'>", unsafe_allow_html=True)
        # Показываем текст; метку изображения заменяем на подпись
        parts = result["text"].split(IMAGE_MARKER)
        for idx, part in enumerate(parts):
            st.markdown(part)
            if idx < len(parts) - 1 and result.get("image") is not None:
                st.image(result["image"], width=420)
        st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# 9. ТОЧКА ВХОДА
# --------------------------------------------------------------------------
if check_password():
    main_app()
