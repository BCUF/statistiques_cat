"""Tkinter interface for choosing and generating Excel exports."""

import json
import os
import traceback
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from config import YEAR, validate_year


ProgressCallback = Callable[[int, int, str], None]


def _settings_file_path() -> Path:
    app_data = os.environ.get("APPDATA")
    settings_root = Path(app_data) if app_data else Path.home() / ".config"
    return settings_root / "Statistiques CAT" / "gui_settings.json"


def _load_gui_settings(settings_file: Path) -> dict[str, object]:
    try:
        with settings_file.open(encoding="utf-8") as file:
            settings = json.load(file)
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError(
            f"Impossible de lire les préférences GUI dans {settings_file}: {error}"
        ) from error

    if not isinstance(settings, dict):
        raise ValueError(
            f"Le fichier de préférences GUI {settings_file} doit contenir un objet JSON."
        )

    for key in ("input_file", "output_folder", "export_choice", "year"):
        if key in settings and not isinstance(settings[key], str):
            raise ValueError(
                f"La préférence GUI « {key} » est invalide dans {settings_file}."
            )

    output_folders = settings.get("output_folders", {})
    if not isinstance(output_folders, dict) or any(
        not isinstance(product, str) or not isinstance(folder, str)
        for product, folder in output_folders.items()
    ):
        raise ValueError(
            f"Les dossiers de sortie enregistrés sont invalides dans {settings_file}."
        )

    return settings


def _save_gui_settings(settings_file: Path, settings: dict[str, object]) -> None:
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    with settings_file.open("w", encoding="utf-8") as file:
        json.dump(settings, file, ensure_ascii=False, indent=2)
        file.write("\n")


def lancer_interface(
    export_choices: dict[str, str],
    generer_exports: Callable[
        [str, str, str, dict[str, str], str, ProgressCallback],
        None,
    ],
    input_file: str,
    output_folder: str,
    default_year: str = YEAR,
) -> None:
    root = tk.Tk()
    root.title("Statistiques CAT")
    root.resizable(False, True)
    executor = ThreadPoolExecutor(max_workers=1)
    settings_file = _settings_file_path()

    frame = ttk.Frame(root, padding=16)
    frame.grid()

    ttk.Label(
        frame,
        text="Que souhaitez-vous générer ?",
        font=("", 11, "bold"),
    ).grid(row=0, column=0, sticky="w", pady=(0, 8))

    output_products = {
        "errors": "Erreurs",
        "valid_fields": "Champs corrects",
        "employees": "Par personne",
        "libraries": "Par bibliothèque",
    }
    try:
        settings = _load_gui_settings(settings_file)
    except ValueError as error:
        messagebox.showwarning(
            "Préférences GUI invalides",
            f"{error}\nLes valeurs par défaut seront utilisées.",
            parent=root,
        )
        settings = {}

    saved_output_folders = settings.get("output_folders", {})
    input_file_var = tk.StringVar(
        value=(
            settings.get("input_file")
            if isinstance(settings.get("input_file"), str)
            else os.path.abspath(input_file)
        )
    )
    output_folder_var = tk.StringVar(
        value=(
            settings.get("output_folder")
            if isinstance(settings.get("output_folder"), str)
            else os.path.abspath(output_folder)
        )
    )
    output_folder_vars = {
        product: tk.StringVar(
            value=(
                saved_output_folders.get(product, "")
                if isinstance(saved_output_folders, dict)
                else ""
            )
        )
        for product in output_products
    }

    ttk.Label(frame, text="Fichier d'entrée :").grid(
        row=1, column=0, sticky="w"
    )
    input_row = ttk.Frame(frame)
    input_row.grid(row=2, column=0, sticky="ew", pady=(2, 8))
    ttk.Entry(input_row, textvariable=input_file_var, width=48).grid(
        row=0, column=0, sticky="ew"
    )
    input_row.columnconfigure(0, weight=1)

    def choisir_dossier(variable: tk.StringVar) -> None:
        initial_dir = variable.get()
        if not os.path.isdir(initial_dir):
            initial_dir = None
        selected_dir = filedialog.askdirectory(
            parent=root,
            initialdir=initial_dir,
            title="Choisir un dossier",
        )
        if selected_dir:
            variable.set(selected_dir)

    def choisir_fichier(variable: tk.StringVar) -> None:
        initial_dir = os.path.dirname(variable.get())
        if not os.path.isdir(initial_dir):
            initial_dir = None
        selected_file = filedialog.askopenfilename(
            parent=root,
            initialdir=initial_dir,
            title="Choisir un fichier",
            filetypes=[("Fichiers Excel", "*.xlsx")]
        )
        if selected_file:
            variable.set(selected_file)

    ttk.Button(
        input_row,
        text="Parcourir…",
        command=lambda: choisir_fichier(input_file_var),
    ).grid(row=0, column=1, padx=(6, 0))

    ttk.Label(frame, text="Année des données :").grid(
        row=3, column=0, sticky="w"
    )
    year_var = tk.StringVar(
        value=(
            settings.get("year")
            if isinstance(settings.get("year"), str)
            else default_year
        )
    )
    ttk.Spinbox(
        frame,
        from_=1000,
        to=9999,
        textvariable=year_var,
        width=8,
    ).grid(row=4, column=0, sticky="w", pady=(2, 8))

    ttk.Label(frame, text="Dossier de sortie par défaut :").grid(
        row=5, column=0, sticky="w"
    )
    output_row = ttk.Frame(frame)
    output_row.grid(row=6, column=0, sticky="ew", pady=(2, 8))
    ttk.Entry(output_row, textvariable=output_folder_var, width=48).grid(
        row=0, column=0, sticky="ew"
    )
    output_row.columnconfigure(0, weight=1)
    ttk.Button(
        output_row,
        text="Parcourir…",
        command=lambda: choisir_dossier(output_folder_var),
    ).grid(row=0, column=1, padx=(6, 0))

    ttk.Label(
        frame,
        text=(
            "Dossiers spécifiques (facultatifs, sinon dossier par défaut) :"
        ),
    ).grid(row=7, column=0, sticky="w", pady=(2, 4))
    for row, (product, label) in enumerate(output_products.items(), start=8):
        product_row = ttk.Frame(frame)
        product_row.grid(row=row, column=0, sticky="ew", pady=2)
        ttk.Label(product_row, text=f"{label} :", width=16).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(
            product_row,
            textvariable=output_folder_vars[product],
            width=38,
        ).grid(row=0, column=1, sticky="ew")
        product_row.columnconfigure(1, weight=1)
        ttk.Button(
            product_row,
            text="Parcourir…",
            command=lambda variable=output_folder_vars[product]: choisir_dossier(
                variable
            ),
        ).grid(row=0, column=2, padx=(6, 0))

    saved_choice = settings.get("export_choice")
    if not isinstance(saved_choice, str) or saved_choice not in export_choices:
        saved_choice = next(iter(export_choices), "1")
    choix_var = tk.StringVar(value=saved_choice)
    first_choice_row = 8 + len(output_products)
    for row, (choix, description) in enumerate(
        export_choices.items(),
        start=first_choice_row,
    ):
        ttk.Radiobutton(
            frame,
            text=description,
            variable=choix_var,
            value=choix,
        ).grid(row=row, column=0, sticky="w", pady=2)

    status = tk.StringVar(value=" ")
    status_row = first_choice_row + len(export_choices)
    ttk.Label(frame, textvariable=status, wraplength=420).grid(
        row=status_row, column=0, sticky="w", pady=(8, 4)
    )
    progress_bar = ttk.Progressbar(
        frame,
        mode="indeterminate",
        length=420,
    )
    progress_bar.grid(
        row=status_row + 1,
        column=0,
        sticky="ew",
        pady=(0, 4),
    )
    busy = {"value": False}

    def sauvegarder_preferences() -> None:
        settings = {
            "input_file": input_file_var.get().strip(),
            "output_folder": output_folder_var.get().strip(),
            "year": year_var.get().strip(),
            "output_folders": {
                product: variable.get().strip()
                for product, variable in output_folder_vars.items()
            },
            "export_choice": choix_var.get(),
        }
        try:
            _save_gui_settings(settings_file, settings)
        except OSError as error:
            messagebox.showerror(
                "Enregistrement des préférences impossible",
                f"Les préférences GUI n'ont pas pu être enregistrées dans "
                f"{settings_file}.\n\n{error}",
                parent=root,
            )

    def afficher_progression(
        completed: int,
        total: int,
        description: str,
    ) -> None:
        if total <= 0:
            if progress_bar["mode"] != "indeterminate":
                progress_bar.stop()
                progress_bar.configure(mode="indeterminate")
                progress_bar.start(10)
            status.set(description)
            return

        if progress_bar["mode"] == "indeterminate":
            progress_bar.stop()
            progress_bar.configure(
                mode="determinate",
                maximum=total,
                value=0,
            )
        else:
            progress_bar.configure(maximum=total)
        progress_bar["value"] = completed
        status.set(f"{description} ({completed}/{total})")

    def signaler_progression(
        completed: int,
        total: int,
        description: str,
    ) -> None:
        root.after(
            0,
            afficher_progression,
            completed,
            total,
            description,
        )

    def lancer() -> None:
        choix = choix_var.get()
        selected_input = input_file_var.get().strip()
        selected_output = output_folder_var.get().strip()
        selected_year = year_var.get().strip()
        selected_output_folders = {
            product: variable.get().strip()
            for product, variable in output_folder_vars.items()
            if variable.get().strip()
        }
        try:
            selected_year = validate_year(selected_year)
        except ValueError as error:
            messagebox.showerror(
                "Année invalide",
                str(error),
                parent=root,
            )
            return

        if not selected_input or not os.path.isfile(selected_input):
            messagebox.showerror(
                "Fichier d'entrée invalide",
                "Veuillez choisir un fichier d'entrée existant.",
                parent=root,
            )
            return
        if not selected_output:
            messagebox.showerror(
                "Dossier de sortie invalide",
                "Veuillez indiquer un dossier de sortie.",
                parent=root,
            )
            return

        sauvegarder_preferences()
        busy["value"] = True
        bouton_lancer.configure(state="disabled")
        bouton_annuler.configure(state="disabled")
        status.set("Génération en cours…")
        progress_bar.configure(mode="indeterminate", value=0)
        progress_bar.start(10)

        future = executor.submit(
            generer_exports,
            choix,
            selected_input,
            selected_output,
            selected_output_folders,
            selected_year,
            signaler_progression,
        )

        def terminer(future) -> None:
            error = future.exception()
            if error is not None:
                details = "".join(traceback.format_exception(error))
                root.after(0, afficher_erreur, error, details)
            else:
                root.after(0, afficher_succes)

        future.add_done_callback(terminer)

    def afficher_erreur(error: Exception, details: str) -> None:
        busy["value"] = False
        progress_bar.stop()
        status.set("La génération a échoué.")
        bouton_lancer.configure(state="normal")
        bouton_annuler.configure(state="normal")
        print(details)
        messagebox.showerror(
            "Erreur pendant la génération",
            f"Les fichiers n'ont pas pu être générés.\n\n{error}",
            parent=root,
        )

    def afficher_succes() -> None:
        progress_bar.stop()
        status.set("Les fichiers demandés ont été générés.")
        messagebox.showinfo(
            "Génération terminée",
            "Les fichiers demandés ont été générés.",
            parent=root,
        )
        root.destroy()

    def fermer() -> None:
        if not busy["value"]:
            sauvegarder_preferences()
            root.destroy()

    boutons = ttk.Frame(frame)
    boutons.grid(row=status_row + 2, column=0, sticky="e", pady=(8, 0))
    bouton_lancer = ttk.Button(boutons, text="Générer", command=lancer)
    bouton_lancer.grid(row=0, column=0, padx=(0, 6))
    bouton_annuler = ttk.Button(boutons, text="Annuler", command=fermer)
    bouton_annuler.grid(row=0, column=1)
    root.protocol("WM_DELETE_WINDOW", fermer)

    root.mainloop()
