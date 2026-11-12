# PDF + CAD Register — Standalone Windows EXE

The GitHub Actions workflow builds a Windows 64-bit one-file EXE and bundles
LibreDWG so a separate ODA File Converter installation is not required.

## Download

1. Open the repository's **Actions** tab.
2. Open the newest successful **Build Windows EXE** run.
3. Download the artifact **PDF-and-CAD-Register-Windows-Standalone**.
4. Extract the artifact ZIP.
5. Run **PDF and CAD Register.exe**.

The package also contains LibreDWG's GPL-3.0-or-later licence and corresponding
source archive, as required for redistribution of the bundled converter.

## Why the app used to be slow and open black command windows

The previous version started the external LibreDWG `dwg2dxf.exe` process once
for every DWG. That caused a visible console flash and repeated process startup
overhead.

The current build:
- Uses Windows no-console startup flags to keep the command window hidden.
- Passes up to 20 distinct-named DWG files to one `dwg2dxf` invocation.
- Splits very large folders into safe-sized batches.
- Handles duplicate base filenames in isolated temporary folders to prevent
  one drawing overwriting another.
- Parses generated DXFs into memory and deletes temporary conversion files.
- Reports batch/file progress and checks for cancellation between batches.

LibreDWG's command-line manual documents multiple `DWGFILES` in one command;
when processing multiple inputs, `-o` is not used and the default DXFs are
written to the current working directory. The application sets that working
directory to a temporary folder.

Batch mode reduces process-launch overhead. LibreDWG still decodes drawings
internally, so complex or large DWG files will still take time.

## DWG conversion

The EXE bundles LibreDWG's `dwg2dxf.exe` and DLLs. It uses LibreDWG first;
ODA File Converter is only an optional fallback. AutoCAD and a separately
installed ODA converter should not be required by this package.

LibreDWG describes its DWG-to-DXF support as approximately 90% coverage.
Some complex, proxy, raster, or unsupported entities may not convert. The
application marks unreadable/incomplete drawings for manual review.

## Test and validation status

The supplied sample PDF has title `DATA CENTER - GROUND FLOOR FORMWORK`,
serial `ST111`, date `27/10/26`, and title-block revision `00`. Its filename
ends in `_01`, so the revision should be marked as a mismatch.

The exact supplied DWG has not yet been validated by this environment using
the final Windows EXE and bundled converter. Test it on a copy of the actual
project folder before relying on the register for submissions.
