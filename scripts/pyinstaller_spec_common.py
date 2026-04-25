from pathlib import Path


PROJECT_ROOT = Path.cwd()
APP_NAME = "售后登记表_v1.7"
ENTRY_SCRIPT = str(PROJECT_ROOT / "dj.py")
ICON_PATH = None

DATA_FILES = [
    (str(PROJECT_ROOT / "dialog_add_store.ui"), "."),
    (str(PROJECT_ROOT / "dialog_store_settings.ui"), "."),
    (str(PROJECT_ROOT / "input_panel.ui"), "."),
    (str(PROJECT_ROOT / "main_window.ui"), "."),
    (str(PROJECT_ROOT / "quick_date_panel.ui"), "."),
    (str(PROJECT_ROOT / "search_panel.ui"), "."),
    (str(PROJECT_ROOT / "table_panel.ui"), "."),
    (str(PROJECT_ROOT / "dopamine_styles.qss"), "."),
    (str(PROJECT_ROOT / "help_dialog.py"), "."),
    (str(PROJECT_ROOT / "icons"), "icons"),
]

HIDDEN_IMPORTS = [
    "PyQt5.uic",
    "matplotlib.backends.backend_qt5agg",
    "markdown.extensions.extra",
]

EXCLUDES = [
    "pandas",
    "xlrd",
    "setuptools",
    "pkg_resources",
    "wheel",
    "pytest",
    "unittest",
    "test",
    "tests",
    "IPython",
    "jupyter_client",
    "jupyter_core",
    "notebook",
    "PIL.ImageQt",
    "Pythonwin",
    "pythonwin",
    "win32com",
    "pywin32",
]

MATPLOTLIB_PRUNE_MARKERS = (
    "matplotlib/mpl-data/sample_data/",
    "matplotlib/mpl-data/fonts/afm/",
    "matplotlib/mpl-data/fonts/pdfcorefonts/",
    "matplotlib/mpl-data/stylelib/",
    "matplotlib/tests/",
)


def _normalize_entry_text(entry):
    parts = []
    for value in entry:
        if isinstance(value, bytes):
            parts.append(value.decode("utf-8", errors="ignore"))
        elif isinstance(value, str):
            parts.append(value)
    return " ".join(parts).replace("\\", "/").lower()


def prune_analysis_datas(entries):
    filtered = []
    for entry in entries:
        entry_text = _normalize_entry_text(entry)
        if any(marker in entry_text for marker in MATPLOTLIB_PRUNE_MARKERS):
            continue
        filtered.append(entry)
    return filtered
