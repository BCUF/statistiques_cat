"""Tkinter interface for choosing and generating Excel exports."""

import os
import traceback
from concurrent.futures import ThreadPoolExecutor
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def lancer_interface(
    export_choices: dict[str, str],
    generer_exports: Callable[[str, str, str], None],
    input_folder: str,
    output_folder: str,
) -> None:
    root = tk.Tk()
    root.title("Statistiques CAT")
    root.resizable(False, True)
    executor = ThreadPoolExecutor(max_workers=1)

    frame = ttk.Frame(root, padding=16)
    frame.grid()

    ttk.Label(
        frame,
        text="Que souhaitez-vous générer ?",
        font=("", 11, "bold"),
    ).grid(row=0, column=0, sticky="w", pady=(0, 8))

    input_folder_var = tk.StringVar(value=os.path.abspath(input_folder))
    output_folder_var = tk.StringVar(value=os.path.abspath(output_folder))

    ttk.Label(frame, text="Dossier des fichiers d'entrée :").grid(
        row=1, column=0, sticky="w"
    )
    input_row = ttk.Frame(frame)
    input_row.grid(row=2, column=0, sticky="ew", pady=(2, 8))
    ttk.Entry(input_row, textvariable=input_folder_var, width=48).grid(
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

    ttk.Button(
        input_row,
        text="Parcourir…",
        command=lambda: choisir_dossier(input_folder_var),
    ).grid(row=0, column=1, padx=(6, 0))

    ttk.Label(frame, text="Dossier des fichiers de sortie :").grid(
        row=3, column=0, sticky="w"
    )
    output_row = ttk.Frame(frame)
    output_row.grid(row=4, column=0, sticky="ew", pady=(2, 8))
    ttk.Entry(output_row, textvariable=output_folder_var, width=48).grid(
        row=0, column=0, sticky="ew"
    )
    output_row.columnconfigure(0, weight=1)
    ttk.Button(
        output_row,
        text="Parcourir…",
        command=lambda: choisir_dossier(output_folder_var),
    ).grid(row=0, column=1, padx=(6, 0))

    choix_var = tk.StringVar(value="1")
    for row, (choix, description) in enumerate(export_choices.items(), start=1):
        ttk.Radiobutton(
            frame,
            text=description,
            variable=choix_var,
            value=choix,
        ).grid(row=row + 4, column=0, sticky="w", pady=2)

    status = tk.StringVar(value=" ")
    ttk.Label(frame, textvariable=status, wraplength=420).grid(
        row=12, column=0, sticky="w", pady=(8, 4)
    )
    busy = {"value": False}

    def lancer() -> None:
        choix = choix_var.get()
        selected_input = input_folder_var.get().strip()
        selected_output = output_folder_var.get().strip()

        if not selected_input or not os.path.isdir(selected_input):
            messagebox.showerror(
                "Dossier d'entrée invalide",
                "Veuillez choisir un dossier d'entrée existant.",
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

        busy["value"] = True
        bouton_lancer.configure(state="disabled")
        bouton_annuler.configure(state="disabled")
        status.set("Génération en cours…")

        future = executor.submit(
            generer_exports,
            choix,
            selected_input,
            selected_output,
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
        status.set("Les fichiers demandés ont été générés.")
        messagebox.showinfo(
            "Génération terminée",
            "Les fichiers demandés ont été générés.",
            parent=root,
        )
        root.destroy()

    def fermer() -> None:
        if not busy["value"]:
            root.destroy()

    boutons = ttk.Frame(frame)
    boutons.grid(row=13, column=0, sticky="e", pady=(8, 0))
    bouton_lancer = ttk.Button(boutons, text="Générer", command=lancer)
    bouton_lancer.grid(row=0, column=0, padx=(0, 6))
    bouton_annuler = ttk.Button(boutons, text="Annuler", command=fermer)
    bouton_annuler.grid(row=0, column=1)
    root.protocol("WM_DELETE_WINDOW", fermer)

    root.mainloop()
