# DealFlowApp.spec
import sys
import glob as _glob
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

SPEC_DIR = Path(SPECPATH)

# NiceGUI web assets (Vue, Quasar, etc.) - MUST be included
nicegui_datas = collect_data_files('nicegui')

# App assets
app_datas = [
    (str(SPEC_DIR / 'rich_editor.js'), '.'),       # same dir as rich_editor.pyc in _MEIPASS
    (str(SPEC_DIR / 'templates'), 'templates'),     # seed templates
]

# pywin32 DLLs (pywintypes3xx.dll, pythoncom3xx.dll)
venv_root = Path(sys.executable).parent.parent
pywin32_sys32 = venv_root / 'Lib' / 'site-packages' / 'pywin32_system32'
pywin32_binaries = [(str(dll), '.') for dll in _glob.glob(str(pywin32_sys32 / '*.dll'))]

hidden_imports = [
    'nicegui.elements', 'nicegui.elements.mixins', 'nicegui.templates',
    'win32com', 'win32com.client', 'win32com.client.dynamic', 'win32com.client.gencache',
    'win32api', 'win32con', 'pywintypes', 'pythoncom',
    'openpyxl.cell._writer', 'openpyxl.styles.stylesheet',
    'asyncio', 'multiprocessing', 'multiprocessing.freeze_support', 'multiprocessing.spawn',
    'uvicorn.logging', 'uvicorn.loops', 'uvicorn.loops.auto',
    'uvicorn.protocols', 'uvicorn.protocols.http', 'uvicorn.protocols.http.auto',
    'uvicorn.protocols.websockets', 'uvicorn.protocols.websockets.auto',
    'uvicorn.lifespan', 'uvicorn.lifespan.on',
    'starlette.routing', 'starlette.middleware', 'starlette.middleware.cors',
    'encodings', 'encodings.utf_8', 'encodings.cp1252',
] + collect_submodules('nicegui')

a = Analysis(
    ['main.py'],
    pathex=[str(SPEC_DIR)],
    binaries=pywin32_binaries,
    datas=nicegui_datas + app_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    runtime_hooks=[],
    excludes=['pytest', 'unittest', 'tkinter', '_tkinter', 'IPython', 'jupyter'],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='SourceThatDeal',
    debug=False,
    strip=False,
    upx=False,        # NEVER use UPX - #1 AV false positive trigger
    console=False,    # no console window; change to True for debugging
    icon=None,        # replace with path to .ico if available
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    strip=False, upx=False,
    name='SourceThatDeal',
)
