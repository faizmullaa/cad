from pathlib import Path

source_path = Path("pdf_cad_register.py")
source = source_path.read_text(encoding="utf-8")
start = source.index("def _open_dwg_as_dxf(")
end = source.index("\ndef read_dwg(", start)

replacement = r'''def _bundled_resource(*parts):
    """Resolve a resource packaged by PyInstaller or stored beside the source."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


def _run_libredwg(converter, path):
    """Convert one DWG silently and read its temporary DXF."""
    import tempfile
    temp_dir = tempfile.mkdtemp(prefix="cad_register_")
    out = os.path.join(temp_dir, os.path.splitext(os.path.basename(path))[0] + ".dxf")

    # dwg2dxf is a console-subsystem program. Without these Windows flags,
    # Windows opens a new black console window for every drawing converted.
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        creationflags = subprocess.CREATE_NO_WINDOW

    try:
        proc = subprocess.run(
            [converter, "-y", "-o", out, path],
            capture_output=True,
            text=True,
            timeout=180,
            cwd=os.path.dirname(converter) or None,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if proc.returncode == 0 and os.path.isfile(out):
            return ezdxf.readfile(out)
        detail = (proc.stderr or proc.stdout or "DWG conversion failed.").strip()[-900:]
        raise RuntimeError(detail)
    finally:
        # The DXF document is loaded in memory; remove the temporary conversion
        # file and directory instead of leaving one behind for every drawing.
        shutil.rmtree(temp_dir, ignore_errors=True)


def _open_dwg_as_dxf(path):
    """Use the bundled LibreDWG engine first; ODA is an optional fallback."""
    bundled_converter = _bundled_resource("libredwg", "dwg2dxf.exe")
    if os.path.isfile(bundled_converter):
        return _run_libredwg(bundled_converter, path)

    converter = shutil.which("dwg2dxf") or shutil.which("dwg2dxf.exe")
    if converter:
        return _run_libredwg(converter, path)

    try:
        from ezdxf.addons import odafc
        if odafc.is_installed():
            return odafc.readfile(path)
    except Exception as exc:
        oda_error = str(exc)
    else:
        oda_error = "No optional system converter was detected."

    raise RuntimeError(
        "The bundled LibreDWG converter is missing. Re-download the complete "
        "application package. " + oda_error
    )
'''
updated_source = source[:start] + replacement + source[end:]
updated_source = updated_source.replace(
    "Could not read this DWG. Check that it is valid and that a DWG converter is installed.",
    "Could not read this DWG. It may be damaged or use unsupported CAD entities."
)
source_path.write_text(updated_source, encoding="utf-8")
compile(updated_source, "pdf_cad_register.py", "exec")
