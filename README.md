# PDF + CAD Register — Windows EXE

This repository builds a Windows desktop executable through GitHub Actions.

## Download the EXE

1. Open the repository's **Actions** tab.
2. Select the latest **Build Windows EXE** run.
3. Wait for the run to finish with a green check.
4. Under **Artifacts**, download **PDF-and-CAD-Register-Windows**.
5. Extract the downloaded ZIP. It contains **PDF and CAD Register.exe**.

The workflow runs on every push to `main` and can also be started manually with **Actions → Build Windows EXE → Run workflow**.

## What the app does

Select a folder and scan its PDF and DWG files (including subfolders) to create an Excel workbook with Documents, Drawings, Shop Drawings and Summary sheets. Rows with mismatches are marked for review; incomplete or unreadable files are marked CHECK.

## Important DWG requirement

The application does not include a DWG conversion engine. For DWG text/title-block extraction, install a compatible DWG-to-DXF converter on the Windows computer that runs the app, such as [ODA File Converter](https://www.opendesign.com/guestfiles/oda_file_converter). AutoCAD is not required. Converter availability and its own licensing/redistribution terms should be checked separately.

If a converter is unavailable or cannot read a file, the app should show the DWG row as CHECK. The EXE build succeeding does not prove DWG extraction works for every DWG.

## Validation note

The app checks filename serial/drawing numbers against values extracted from title blocks. It also compares the filename revision suffix (for example, `_01`) with the revision extracted from the drawing title block. Both values are displayed separately in the Drawings and Shop Drawings sheets.

The supplied PDF test sample has filename suffix `_01` and title-block revision `00`; the revised code should mark this as MISMATCH. Confirm against project drawings before operational use.

## Build details

GitHub Actions uses a Windows runner and Python 3.11 to restore the compressed Python source, install dependencies and build a one-file GUI executable using PyInstaller. The application source payload is stored as `build_payload.py.xz.b64` and decoded during the workflow.
