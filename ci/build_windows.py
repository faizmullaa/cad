from pathlib import Path
import sys
from PyInstaller.__main__ import run

vendor = Path(sys.argv[1]).resolve()
exe = vendor / "dwg2dxf.exe"
if not exe.is_file():
    raise SystemExit(f"LibreDWG executable missing: {exe}")

binaries = [exe, *sorted(p for p in vendor.glob("*.dll") if p.name.casefold() != "_libredwg.dll")]
args = [
    "--noconfirm", "--clean", "--onefile", "--windowed",
    "--name", "PDF and CAD Register",
    "--collect-all", "ezdxf",
    "--collect-all", "pymupdf",
    "--collect-all", "sv_ttk",
    "--hidden-import", "openpyxl",
    "--hidden-import", "pymupdf",
    "pdf_cad_register.py",
]
for item in binaries:
    args.extend(["--add-binary", f"{item};libredwg"])
print("Bundling LibreDWG files:", ", ".join(p.name for p in binaries))
run(args)
