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

## DWG conversion

The EXE bundles LibreDWG's `dwg2dxf.exe` and DLLs. It uses LibreDWG first;
ODA File Converter is only an optional fallback. AutoCAD and a separately
installed ODA converter should not be required by this package.

LibreDWG currently documents approximately 90% coverage. Some complex,
proxy, raster, or unsupported entities may not convert. The application must
mark unreadable/incomplete drawings for manual review.

## Test and validation status

The supplied sample PDF has title `DATA CENTER - GROUND FLOOR FORMWORK`,
serial `ST111`, date `27/10/26`, and title-block revision `00`. Its filename
ends in `_01`, so the revision should be marked as a mismatch.

The supplied DWG was previously detected but not parsed because no converter
was available in the test environment. The new Windows workflow verifies the
bundled converter itself (`dwg2dxf --version`) but cannot claim the supplied
project DWG was parsed until that exact file has been tested with this EXE.
