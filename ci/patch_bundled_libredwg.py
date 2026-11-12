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
    """Convert one DWG to DXF using LibreDWG, then read the converted file."""
    import tempfile
    temp_dir = tempfile.mkdtemp(prefix="cad_register_")
    out = os.path.join(temp_dir, os.path.splitext(os.path.basename(path))[0] + ".dxf")
    proc = subprocess.run(
        [converter, "-y", "-o", out, path],
        capture_output=True,
        text=True,
        timeout=180,
        cwd=os.path.dirname(converter) or None,
    )
    if proc.returncode == 0 and os.path.isfile(out):
        return ezdxf.readfile(out)
    detail = (proc.stderr or proc.stdout or "DWG conversion failed.").strip()[-900:]
    raise RuntimeError(detail)


def _open_dwg_as_dxf(path):
    """Use the bundled LibreDWG engine first; ODA is an optional fallback."""
    # In the one-file EXE, PyInstaller extracts this folder into _MEIPASS.
    # Users do not need to install a separate DWG converter.
    bundled_converter = _bundled_resource("libredwg", "dwg2dxf.exe")
    if os.path.isfile(bundled_converter):
        return _run_libredwg(bundled_converter, path)

    # Developer/source-run fallback for an already installed LibreDWG.
    converter = shutil.which("dwg2dxf") or shutil.which("dwg2dxf.exe")
    if converter:
        return _run_libredwg(converter, path)

    # ODA remains optional for environments that already have it installed.
    try:
        from ezdxf.addons import odafc
        if odafc.is_installed():
            return odafc.readfile(path)
    except Exception as exc:
        oda_error = str(exc)
    else:
        oda_error = "ODA File Converter was not detected (it is optional)."

    raise RuntimeError(
        "The bundled LibreDWG converter is missing and no optional system "
        "converter was found. Re-download the complete application package. "
        + oda_error
    )
'''
source_path.write_text(source[:start] + replacement + source[end:], encoding="utf-8")
compile(source_path.read_text(encoding="utf-8"), str(source_path), "exec")
