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

Running `uv run main.py` without arguments displays the command-line help.

To launch the optional graphical interface do `uv run main.py --gui`. It offers the same export choices as the text menu and lets you select the input and output folders.

For command-line exports, provide the full path to the input workbook and optionally choose an output directory and product:

```powershell
uv run main.py --input-file "C:\Data\champs_990_993_all_20261112.xlsx" --output-dir "C:\Data\Exports" --product all
```

Options are `-i, --input-file` (required for command-line exports), `-o, --output-dir` (defaults to the configured export folder), and `-p, --product` (one or more of `errors`, `valid_fields`, `employees`, `libraries`, or `all`; defaults to `all`). Separate multiple products with spaces, for example `--product errors valid_fields`. These choices generate the errors workbook, correct-fields workbook, per-employee workbooks, per-library workbooks, or all products, respectively.

## Create a release
Push a version tag such as `v0.1.0` to trigger the GitHub Actions release workflow:

```powershell
git tag v0.1.0
git push origin v0.1.0
```

The workflow builds a Windows x64 executable and attaches a ZIP to the GitHub Release for that tag. The ZIP includes `run.cmd` to launch the graphical interface. After extracting it, put the source workbook in a `Données` folder next to `statistiques-cat.exe`, using the filename pattern `champs_993_all_YYYYMMDD.xlsx`.
