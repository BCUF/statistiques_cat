# statistiques_cat
Collect, process, and extract cataloging performance statistics from raw Analytics data.

## Install
This program use uv as package manager (https://github.com/astral-sh/uv).
To install on windows do 
```
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

To intall the dependency on your computer do `uv sync`

## Run

To run the software in the dev environement do `uv run main.py`

## Create a release
Push a version tag such as `v0.1.0` to trigger the GitHub Actions release workflow:

```powershell
git tag v0.1.0
git push origin v0.1.0
```

The workflow builds a Windows x64 executable and attaches a ZIP to the GitHub Release for that tag. After extracting the ZIP, put the source workbook in a `Données` folder next to `statistiques-cat.exe`, using the filename pattern `champs_993_all_YYYYMMDD.xlsx`.
