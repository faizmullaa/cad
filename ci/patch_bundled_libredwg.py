from pathlib import Path

source_path = Path("pdf_cad_register.py")
source = source_path.read_text(encoding="utf-8")

# Replace the single-DWG converter invocation with chunked native CLI batching.
start = source.index("def _run_libredwg(converter, path):")
end = source.index("\ndef _open_dwg_as_dxf(", start)
replacement = r'''def _windows_no_console_kwargs():
    """Prevent console-subsystem helper processes from flashing a terminal."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": subprocess.CREATE_NO_WINDOW,
    }


def _run_libredwg_batch(converter, paths, progress=None, cancel_event=None):
    """Convert many DWGs in bounded dwg2dxf invocations and parse DXFs in memory.

    LibreDWG dwg2dxf accepts multiple DWGFILES in one process. Its -o option
    is intentionally not used for a multi-file call; default DXFs are written
    to the process working directory.
    """
    import tempfile
    from collections import defaultdict

    results = {}
    by_stem = defaultdict(list)
    for path in paths:
        by_stem[os.path.splitext(os.path.basename(path))[0].casefold()].append(path)

    # Duplicate basenames would overwrite each other's output in one directory.
    # Keep those in separate temp subfolders; all distinct basenames are batched.
    unique_paths = [items[0] for items in by_stem.values() if len(items) == 1]
    duplicate_groups = [items for items in by_stem.values() if len(items) > 1]

    batches, current, chars = [], [], 0
    for path in unique_paths:
        size = len(path) + 4
        if current and (len(current) >= 40 or chars + size > 22000):
            batches.append(current)
            current, chars = [], 0
        current.append(path)
        chars += size
    if current:
        batches.append(current)

    total_jobs = len(batches) + sum(len(group) for group in duplicate_groups)
    job_index = 0
    completed_files = 0

    with tempfile.TemporaryDirectory(prefix="cad_register_batch_") as temp_root:
        def convert_group(group, output_dir, use_batch):
            command = [converter, "-y"]
            if use_batch:
                command.extend(group)
            else:
                # The output folder is isolated, so a basename collision is safe.
                command.append(group[0])

            try:
                proc = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=max(180, 90 * len(group)),
                    cwd=output_dir,
                    **_windows_no_console_kwargs(),
                )
                diagnostic = (proc.stderr or proc.stdout or "").strip()[-700:]
            except Exception as exc:
                diagnostic = str(exc)
                proc = None

            for original in group:
                expected = os.path.join(
                    output_dir,
                    os.path.splitext(os.path.basename(original))[0] + ".dxf",
                )
                if not os.path.isfile(expected):
                    # Tolerate converters that choose a different extension case.
                    wanted = os.path.basename(expected).casefold()
                    try:
                        expected = next(
                            os.path.join(output_dir, name)
                            for name in os.listdir(output_dir)
                            if name.casefold() == wanted
                        )
                    except (StopIteration, OSError):
                        results[original] = (
                            None,
                            diagnostic or "LibreDWG did not produce a DXF for this file.",
                        )
                        continue
                try:
                    results[original] = (ezdxf.readfile(expected), "")
                except Exception as exc:
                    results[original] = (None, f"DXF parse failed: {str(exc)[:350]}")

        for batch_number, group in enumerate(batches, 1):
            if cancel_event and cancel_event.is_set():
                raise Cancelled
            job_index += 1
            if progress:
                progress(completed_files, len(paths),
                         f"Converting DWG batch {batch_number}/{len(batches)} ({len(group)} files)")
            convert_group(group, temp_root, use_batch=True)
            completed_files += len(group)
            if progress:
                progress(completed_files, len(paths),
                         f"Converted {completed_files} of {len(paths)} DWG files")

        for duplicate_number, group in enumerate(duplicate_groups, 1):
            for original in group:
                if cancel_event and cancel_event.is_set():
                    raise Cancelled
                job_index += 1
                if progress:
                    progress(completed_files, len(paths),
                             f"Converting duplicate-name DWG {job_index}/{total_jobs}")
                output_dir = os.path.join(temp_root, f"duplicate_{duplicate_number}_{job_index}")
                os.makedirs(output_dir, exist_ok=True)
                convert_group([original], output_dir, use_batch=False)
                completed_files += 1
                if progress:
                    progress(completed_files, len(paths),
                             f"Converted {completed_files} of {len(paths)} DWG files")

    return results


def _run_libredwg(converter, path):
    """Compatibility wrapper for one file, using the same hidden batch engine."""
    result = _run_libredwg_batch(converter, [path])
    doc, error = result[path]
    if error:
        raise RuntimeError(error)
    return doc

'''
source = source[:start] + replacement + source[end:]

# Allow the DWG parser to accept an already-converted in-memory DXF document.
source = source.replace(
    "def read_dwg(path):\n    name = os.path.basename(path).upper()\n    try:\n        doc = _open_dwg_as_dxf(path)\n        words = _words_from_dxf(doc)",
    "def read_dwg(path, converted_doc=None, conversion_error=None):\n    name = os.path.basename(path).upper()\n    if conversion_error:\n        return {\"error\": f\"Could not convert this DWG: {conversion_error[:350]}\"}\n    try:\n        doc = converted_doc if converted_doc is not None else _open_dwg_as_dxf(path)\n        words = _words_from_dxf(doc)"
)

# Replace make_register so all DWGs are converted in batches before per-file parsing.
start = source.index("def make_register(folder, progress=None, recursive=True):")
end = source.index("\ndef summary_text(", start)
make_register = r'''def make_register(folder, progress=None, recursive=True, cancel_event=None):
    """Batch-convert DWGs, parse PDFs/DXFs, and write the Excel register."""
    paths = find_input_files(folder, recursive)
    results = {kind: [] for kind, *_ in SECTIONS}
    row_builders = {kind: build for kind, _, _, build, _ in SECTIONS}

    dwg_paths = [p for p in paths if p.lower().endswith(".dwg")]
    converted = {}
    if dwg_paths:
        converter = _bundled_resource("libredwg", "dwg2dxf.exe")
        if not os.path.isfile(converter):
            converter = shutil.which("dwg2dxf") or shutil.which("dwg2dxf.exe")
        if converter:
            converted = _run_libredwg_batch(
                converter, dwg_paths,
                progress=(lambda i, n, label: progress(i, len(paths) + len(dwg_paths), label))
                if progress else None,
                cancel_event=cancel_event,
            )
        else:
            converted = {p: (None, "No DWG converter is available.") for p in dwg_paths}

    total_units = len(paths) + len(dwg_paths)
    for i, path in enumerate(paths, 1):
        if cancel_event and cancel_event.is_set():
            raise Cancelled
        rel = os.path.relpath(path, folder)
        if progress:
            progress(len(dwg_paths) + i, total_units, rel)
        try:
            if path.lower().endswith(".pdf"):
                info = read_pdf(path)
            else:
                doc, error = converted.get(path, (None, "DWG conversion result missing."))
                info = read_dwg(path, converted_doc=doc, conversion_error=error)
        except Exception:
            ext = os.path.splitext(path)[1].upper()
            info = {"error": f"Could not open the {ext} file (damaged, invalid, or unsupported) - check manually."}
        name = os.path.basename(path).upper()
        kind = info.get("kind") or ("drawing" if "-DWG-" in name else "shop_drawing" if "-SDW-" in name
                                    else "document")
        results[kind].append(row_builders[kind](rel, info))

    out = os.path.join(folder, f"PDF CAD Register {datetime.datetime.now():%Y-%m-%d %H%M%S}.xlsx")
    write_excel(out, folder, results)
    return out, results

'''
source = source[:start] + make_register + source[end:]

# Pass the cancellation event from the UI worker into the register pipeline.
source = source.replace(
    "self.events.put((\"done\", folder, *make_register(folder, progress, recursive)))",
    "self.events.put((\"done\", folder, *make_register(folder, progress, recursive, self.stop)))"
)

source_path.write_text(source, encoding="utf-8")
compile(source, str(source_path), "exec")
print("Patched source: batched DWG conversion, no-console child process, cancellation checks.")
