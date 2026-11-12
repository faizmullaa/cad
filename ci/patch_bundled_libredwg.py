from pathlib import Path

source_path = Path("pdf_cad_register.py")
source = source_path.read_text(encoding="utf-8")


def replace_between(text, start_marker, end_marker, replacement):
    start = text.index(start_marker)
    end = text.index(end_marker, start)
    return text[:start] + replacement.rstrip() + "\n\n\n" + text[end:]


dwg_block = r'''
def _bundled_converter_candidates():
    """Return likely locations for the converter embedded by PyInstaller."""
    roots = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(meipass)
    roots.append(os.path.dirname(os.path.abspath(__file__)))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    if exe_dir not in roots:
        roots.append(exe_dir)

    candidates = []
    for root in roots:
        candidates.extend([
            os.path.join(root, "libredwg", "dwg2dxf.exe"),
            os.path.join(root, "dwg2dxf.exe"),
            os.path.join(root, "libredwg", "bin", "dwg2dxf.exe"),
            os.path.join(root, "bin", "dwg2dxf.exe"),
        ])
    # Remove duplicates but keep search order.
    return list(dict.fromkeys(os.path.normpath(p) for p in candidates))


def _bundled_converter_path():
    """Find only the converter shipped inside this application; no ODA lookup."""
    for candidate in _bundled_converter_candidates():
        if os.path.isfile(candidate):
            return candidate
    return None


def _hidden_subprocess_options():
    """Prevent child console programs from opening a command window on Windows."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": subprocess.CREATE_NO_WINDOW,
    }


def _check_bundled_converter():
    """Return (path, version output); fail with exact embedded search paths."""
    converter = _bundled_converter_path()
    if not converter:
        searched = "\n".join(_bundled_converter_candidates())
        raise RuntimeError(
            "The LibreDWG converter is not embedded in this EXE. Searched:\n" + searched
        )
    proc = subprocess.run(
        [converter, "--version"],
        capture_output=True,
        text=True,
        timeout=30,
        cwd=os.path.dirname(converter),
        **_hidden_subprocess_options(),
    )
    output = (proc.stdout or proc.stderr or "").strip()
    if proc.returncode != 0:
        raise RuntimeError(
            f"The embedded LibreDWG converter was found but could not start "
            f"(exit {proc.returncode}): {output[-500:]}"
        )
    return converter, output


def _run_libredwg_batch(converter, paths, progress=None, cancel_event=None):
    """Run LibreDWG once per batch, then read each produced DXF."""
    import tempfile
    from collections import defaultdict

    results = {}
    by_stem = defaultdict(list)
    for path in paths:
        stem = os.path.splitext(os.path.basename(path))[0].casefold()
        by_stem[stem].append(path)

    unique_paths = [group[0] for group in by_stem.values() if len(group) == 1]
    duplicate_groups = [group for group in by_stem.values() if len(group) > 1]

    batches = []
    current = []
    current_chars = 0
    for path in unique_paths:
        addition = len(path) + 4
        if current and (len(current) >= 20 or current_chars + addition > 18000):
            batches.append(current)
            current, current_chars = [], 0
        current.append(path)
        current_chars += addition
    if current:
        batches.append(current)

    completed = 0
    launch_count = len(batches) + sum(len(group) for group in duplicate_groups)

    with tempfile.TemporaryDirectory(prefix="pdf_cad_register_") as work_root:
        def convert_group(group, output_dir):
            command = [converter, "-y", *group]
            try:
                proc = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=max(180, 90 * len(group)),
                    cwd=output_dir,
                    **_hidden_subprocess_options(),
                )
                diagnostic = (proc.stderr or proc.stdout or "").strip()[-900:]
                if proc.returncode and not diagnostic:
                    diagnostic = f"LibreDWG exited with code {proc.returncode}."
            except Exception as exc:
                diagnostic = f"{type(exc).__name__}: {exc}"

            for original in group:
                expected_name = os.path.splitext(os.path.basename(original))[0] + ".dxf"
                try:
                    actual = next(
                        name for name in os.listdir(output_dir)
                        if name.casefold() == expected_name.casefold()
                    )
                    dxf_path = os.path.join(output_dir, actual)
                    results[original] = (ezdxf.readfile(dxf_path), "")
                except StopIteration:
                    results[original] = (
                        None,
                        diagnostic or "LibreDWG did not create the expected DXF output.",
                    )
                except Exception as exc:
                    results[original] = (
                        None,
                        f"Could not parse generated DXF: {type(exc).__name__}: {str(exc)[:300]}",
                    )

        for batch_no, group in enumerate(batches, 1):
            if cancel_event and cancel_event.is_set():
                raise Cancelled
            if progress:
                progress(completed, len(paths),
                         f"Converting batch {batch_no}/{len(batches)} ({len(group)} DWGs)")
            convert_group(group, work_root)
            completed += len(group)
            if progress:
                progress(completed, len(paths), f"Converted {completed}/{len(paths)} DWGs")

        # Same-named DWGs in different directories must not overwrite one another.
        for dup_no, group in enumerate(duplicate_groups, 1):
            for original in group:
                if cancel_event and cancel_event.is_set():
                    raise Cancelled
                if progress:
                    progress(completed, len(paths),
                             f"Converting duplicate-name drawing {completed + 1}/{len(paths)}")
                output_dir = os.path.join(work_root, f"duplicate_{dup_no}_{completed}")
                os.makedirs(output_dir, exist_ok=True)
                convert_group([original], output_dir)
                completed += 1
                if progress:
                    progress(completed, len(paths), f"Converted {completed}/{len(paths)} DWGs")

    return results


def _open_dwg_as_dxf(path):
    """Convert a single DWG using the LibreDWG executable bundled in this EXE."""
    import tempfile
    converter = _bundled_converter_path()
    if not converter:
        searched = "; ".join(_bundled_converter_candidates())
        raise RuntimeError("Bundled LibreDWG converter missing. Searched: " + searched)

    with tempfile.TemporaryDirectory(prefix="pdf_cad_register_single_") as output_dir:
        proc = subprocess.run(
            [converter, "-y", path],
            capture_output=True,
            text=True,
            timeout=180,
            cwd=output_dir,
            **_hidden_subprocess_options(),
        )
        expected = os.path.splitext(os.path.basename(path))[0] + ".dxf"
        dxf_path = os.path.join(output_dir, expected)
        if proc.returncode != 0 or not os.path.isfile(dxf_path):
            detail = (proc.stderr or proc.stdout or "").strip()[-800:]
            raise RuntimeError(detail or f"LibreDWG failed with exit code {proc.returncode}.")
        return ezdxf.readfile(dxf_path)


def read_dwg(path, converted_doc=None, conversion_error=None):
    name = os.path.basename(path).upper()
    if conversion_error:
        return {"error": f"LibreDWG conversion failed: {str(conversion_error)[:350]}"}
    try:
        doc = converted_doc if converted_doc is not None else _open_dwg_as_dxf(path)
        words = _words_from_dxf(doc)
    except Exception as exc:
        return {"error": f"Could not read this DWG using the bundled LibreDWG converter: {type(exc).__name__}: {str(exc)[:350]}"}
    if not words:
        return {"error": "No readable text entities found in this DWG. Text may be embedded in an unsupported proxy object, raster image, or external reference."}

    block = find_title_block(words)
    has_subject = bool(find_labels(words, "Subject"))
    attrs = _attribute_values(doc)
    is_shop = "-SDW-" in name
    is_drawing = "-DWG-" in name or bool(_first_attr(attrs, "Drawing Title", "Dwg Title", "Serial Number", "Serial No"))

    if is_shop or (block and block[2] == "shop_drawing"):
        info = read_drawing(words, *block) if block else {"kind": "shop_drawing", "title": "", "number": "", "rev": "", "date": ""}
        info["title"] = info.get("title") or _first_attr(attrs, "Drawing Title", "Dwg Title", "Title")
        info["number"] = info.get("number") or _first_attr(attrs, "Drawing Number", "Drawing No", "Dwg Number")
        info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
        info["date"] = info.get("date") or _first_attr(attrs, "Date", "Drawing Date")
        return info

    if is_drawing or block:
        info = read_drawing(words, *block) if block else {"kind": "drawing", "title": "", "number": "", "rev": "", "date": ""}
        info["title"] = info.get("title") or _first_attr(attrs, "Drawing Title", "Dwg Title", "Title")
        info["number"] = info.get("number") or _first_attr(attrs, "Serial Number", "Serial No", "Drawing Number", "Drawing No")
        info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
        info["date"] = info.get("date") or _first_attr(attrs, "Date", "Drawing Date")
        return info

    info = read_document(words) if (has_subject or any(DOC_NO_RE.search(w.text) for w in words)) else {"kind": "document", "doc_no": "", "subject": "", "rev": "", "date": ""}
    info["doc_no"] = info.get("doc_no") or _first_attr(attrs, "Document Number", "Document No", "Doc No", "Doc Number")
    info["subject"] = info.get("subject") or _first_attr(attrs, "Subject", "Document Subject", "Document Title")
    info["rev"] = info.get("rev") or _first_attr(attrs, "Revision", "Rev No", "Rev")
    info["date"] = info.get("date") or _first_attr(attrs, "Submittal Date", "Date")
    return info
'''
source = replace_between(
    source,
    "def _open_dwg_as_dxf(path):",
    "def serial_from_file_name(stem):",
    dwg_block,
)

make_register = r'''
def make_register(folder, progress=None, recursive=True, cancel_event=None):
    """Batch-convert DWGs, then parse all files and write the Excel register."""
    paths = find_input_files(folder, recursive)
    results = {kind: [] for kind, *_ in SECTIONS}
    row_builders = {kind: build for kind, _, _, build, _ in SECTIONS}

    dwg_paths = [path for path in paths if path.lower().endswith(".dwg")]
    converted = {}
    if dwg_paths:
        converter = _bundled_converter_path()
        if converter:
            converted = _run_libredwg_batch(
                converter,
                dwg_paths,
                progress=progress,
                cancel_event=cancel_event,
            )
        else:
            searched = "; ".join(_bundled_converter_candidates())
            message = "Bundled LibreDWG converter not found. Searched: " + searched
            converted = {path: (None, message) for path in dwg_paths}

    if progress and dwg_paths:
        progress(len(dwg_paths), len(dwg_paths) + len(paths), "DWG conversion finished; reading files")

    for index, path in enumerate(paths, 1):
        if cancel_event and cancel_event.is_set():
            raise Cancelled
        rel = os.path.relpath(path, folder)
        if progress:
            progress(len(dwg_paths) + index, len(dwg_paths) + len(paths), rel)
        try:
            if path.lower().endswith(".pdf"):
                info = read_pdf(path)
            else:
                doc, error = converted.get(path, (None, "No conversion result was returned."))
                info = read_dwg(path, converted_doc=doc, conversion_error=error)
        except Exception as exc:
            ext = os.path.splitext(path)[1].upper()
            info = {"error": f"Could not open {ext} file: {type(exc).__name__}: {str(exc)[:300]}"}

        name = os.path.basename(path).upper()
        kind = info.get("kind") or (
            "drawing" if "-DWG-" in name else
            "shop_drawing" if "-SDW-" in name else
            "document"
        )
        results[kind].append(row_builders[kind](rel, info))

    out = os.path.join(
        folder,
        f"PDF CAD Register {datetime.datetime.now():%Y-%m-%d %H%M%S}.xlsx",
    )
    write_excel(out, folder, results)
    return out, results
'''
source = replace_between(
    source,
    "def make_register(folder, progress=None, recursive=True):",
    "def summary_text(",
    make_register,
)

# Add an explicit packaged-runtime diagnostic used by CI and support.
main_marker = "def main():\n"
if main_marker not in source:
    raise SystemExit("Could not locate main() entry point.")
source = source.replace(
    main_marker,
    main_marker + '''    if len(sys.argv) > 1 and sys.argv[1] == "--check-converter":
        try:
            converter, version = _check_bundled_converter()
            print("BUNDLED_CONVERTER_OK")
            print("PATH=" + converter)
            print("VERSION=" + version)
            return
        except Exception as exc:
            print("BUNDLED_CONVERTER_ERROR: " + str(exc))
            sys.exit(2)
''',
    1,
)

# The worker must pass its cancellation flag down to the batch converter.
source = source.replace(
    'make_register(folder, progress, recursive)))',
    'make_register(folder, progress, recursive, self.stop)))',
    1,
)

if "No compatible DWG converter was found" in source or "ODA File Converter was not detected" in source:
    raise SystemExit("Old ODA/PATH fallback text remains after patch.")

compile(source, str(source_path), "exec")
source_path.write_text(source, encoding="utf-8")
print("DWG code patched: bundled path only, no ODA/PATH fallback, runtime self-check, batch processing.")
