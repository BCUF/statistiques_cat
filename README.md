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
The GUI saves the selected input file, export choice, and output folders to `%APPDATA%\Statistiques CAT\gui_settings.json` and restores them on the next launch. During generation, a progress bar shows the data-processing phase and each completed output workbook.

For command-line exports, provide the full path to the input workbook and optionally choose a shared output directory, product-specific output directories, and products:

```powershell
uv run main.py --input-file "C:\Data\champs_990_993_all_20261112.xlsx" --output-dir "C:\Data\Exports" --errors-output-dir "C:\Data\Errors" --valid-fields-output-dir "C:\Data\Correct" --employees-output-dir "C:\Data\Employees" --libraries-output-dir "C:\Data\Libraries" --product all
```

Options are `-i, --input-file` (required for command-line exports), `-o, --output-dir` (shared destination; defaults to the configured export folder), and `-p, --product` (one or more of `errors`, `valid_fields`, `employees`, `libraries`, or `all`; defaults to `all`). Use `--errors-output-dir`, `--valid-fields-output-dir`, `--employees-output-dir`, and `--libraries-output-dir` to override the destination for an individual product. If an individual destination is omitted, that product uses `--output-dir`. Separate multiple products with spaces, for example `--product errors valid_fields`. These choices generate the errors workbook, correct-fields workbook, per-employee workbooks, per-library workbooks, or all products, respectively. In the GUI, enter an optional specific folder for each product; products without a specific folder use the default output folder.

## Create a release
Push a version tag such as `v0.1.0` to trigger the GitHub Actions release workflow:

```powershell
git tag v0.1.0
git push origin v0.1.0
```

The workflow builds a Windows x64 executable and attaches a ZIP to the GitHub Release for that tag. The ZIP includes `run.cmd` to launch the graphical interface. After extracting it, put the source workbook in a `Données` folder next to `statistiques-cat.exe`, using the filename pattern `champs_993_all_YYYYMMDD.xlsx`.
