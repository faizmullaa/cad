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

## Why earlier builds opened black command windows

DWG conversion is performed by LibreDWG's separate `dwg2dxf.exe` console
program. The earlier application launched that program once per DWG without
Windows' no-console flags, so Windows showed a black command window for every
drawing and closed it when conversion ended.

The current build uses Windows `CREATE_NO_WINDOW` and hidden startup flags.
The converter should run silently in the background. Temporary DXF files are
removed after each drawing is parsed, rather than being left in the temp folder.

## DWG conversion

The EXE bundles LibreDWG's `dwg2dxf.exe` and DLLs. It uses LibreDWG first;
ODA File Converter is only an optional fallback. AutoCAD and a separately
installed ODA converter should not be required by this package.

LibreDWG currently documents approximately 90% coverage. Some complex,
proxy, raster, or unsupported entities may not convert. The application must
mark unreadable/incomplete drawings for manual review.

## Performance notes

PDF text extraction reads only the first three pages. DWG files require a
binary CAD conversion before text can be extracted, and the application
currently converts them sequentially, one drawing at a time. Large DWGs or
folders containing many DWGs will therefore take longer than PDF-only folders.
The new build removes the visible console-window flashes and cleans temporary
conversion output; it does not make the DWG decoder itself instantaneous.

## Test and validation status

The supplied sample PDF has title `DATA CENTER - GROUND FLOOR FORMWORK`,
serial `ST111`, date `27/10/26`, and title-block revision `00`. Its filename
ends in `_01`, so the revision should be marked as a mismatch.

The supplied DWG was previously detected but not parsed because no converter
was available in the test environment. The new Windows workflow verifies the
bundled converter itself (`dwg2dxf --version`) but cannot claim the supplied
project DWG was parsed until that exact file has been tested with this EXE.
