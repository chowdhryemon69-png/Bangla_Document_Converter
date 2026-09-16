from pathlib import Path

from PyInstaller.building.build_main import Analysis, COLLECT, EXE, PYZ
from PyInstaller.building.datastruct import TOC
from PyInstaller.utils.hooks import collect_submodules

project_dir = Path(SPECPATH)
node_modules_dir = project_dir / "node_modules" / "bijoy-unicode-converter"
node_executable = Path(r"C:\Program Files\nodejs\node.exe")

runtime_data = [
    (str(project_dir / "bijoy_bridge.mjs"), "runtime"),
    (str(project_dir / "packaging" / "runtime" / "LICENSE.node.txt"), "runtime"),
]
for file_path in node_modules_dir.rglob("*"):
    if file_path.is_file():
        destination = Path("runtime") / "node_modules" / "bijoy-unicode-converter" / file_path.relative_to(node_modules_dir)
        runtime_data.append((str(file_path), str(destination.parent)))

runtime_binaries = [(str(node_executable), "runtime")]

hiddenimports = collect_submodules("docx")

a = Analysis(
    [str(project_dir / "gui.py")],
    pathex=[str(project_dir)],
    binaries=runtime_binaries,
    datas=runtime_data,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="BanglaDocumentConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="BanglaDocumentConverter",
)
