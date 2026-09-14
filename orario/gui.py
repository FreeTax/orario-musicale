"""Finestra del programma: apri file → fogli a griglia modificabili → Calcola → Excel + PDF."""

from __future__ import annotations

import json
import os
import platform
import queue
import re
import shutil
import subprocess
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from openpyxl.utils import get_column_letter
from tksheet import Sheet

from .costanti import (
    ETICHETTA_ACCOMP, FOGLI_DATI, FOGLIO_ABBINAMENTI, FOGLIO_DOCENTI, FOGLIO_GRUPPI, FOGLIO_IMPEGNI,
    FOGLIO_LMI, FOGLIO_STUDENTI, GIORNI_LUNGHI, NOMI_TIPO, ORE,
    N_ORE, ORE, ORE_FINE, ORE_LABEL, TIPO_ACCOMP, TIPO_LMC, TIPO_STRUM1, fascia,
)
from .lettura import (
    Tabella, aggiorna_struttura, aggiorna_totali_docenti, applica_trasporti, costruisci_dati, leggi_orario_precedente, leggi_tabelle, leggi_trasporti,
    salva_orario_nel_file, salva_tabelle, salva_trasporti,
)
from .modello import Orario, Problema, ProblemiError
from .template import ISTRUZIONI, crea_nuovo_file

VERSIONE = "1.0 (settembre 2026)"

IMPOSTAZIONI = Path.home() / ".orario_musicale.json"
# Helvetica non esiste su Windows: Segoe UI è il carattere di sistema
CARATTERE = "Segoe UI" if platform.system() == "Windows" else "Helvetica"
BINDINGS = (
    # selezione con mouse e trackpad
    "single_select", "drag_select", "ctrl_select", "toggle_select", "row_select", "column_select", "select_all",
    # spostamento con la tastiera, come in Excel
    "arrowkeys", "tab", "next", "prior", "row_start_bindings", "table_start_bindings",
    # copia, incolla, taglia, annulla, ripeti, cancella
    "copy", "cut", "paste", "delete", "undo", "redo",
    # modifica delle celle, ricerca e sostituzione
    "edit_cell", "find", "replace",
    # menu con il tasto destro, inserimento ed eliminazione di righe
    "right_click_popup_menu", "rc_select", "rc_insert_row", "rc_delete_row",
    # ridimensionamento di colonne e righe con il mouse
    "column_width_resize", "double_click_column_resize", "row_height_resize", "double_click_row_resize",
    # spostare le righe trascinandole, ordinare per colonna
    "move_rows", "sort_rows",
)
OPZIONI_FOGLIO = dict(
    paste_can_expand_x=False,        # le colonne sono fisse: incollando non se ne aggiungono
    paste_can_expand_y=True,         # incollando più righe di quante ce ne sono, le righe si aggiungono
    expand_sheet_if_paste_too_big=True,
    paste_insert_row_limit=5000,
    edit_cell_return="down",         # Invio scende di una riga, come in Excel
    edit_cell_tab="right",           # Tab passa alla colonna successiva
    max_undos=100,
    set_cell_sizes_on_zoom=True,
)


def _carica_impostazioni() -> dict:
    try:
        return json.loads(IMPOSTAZIONI.read_text())
    except Exception:
        return {}


def _salva_impostazioni(d: dict) -> None:
    try:
        IMPOSTAZIONI.write_text(json.dumps(d))
    except Exception:
        pass


def apri_cartella(p: Path) -> None:
    try:
        if platform.system() == "Darwin":
            subprocess.Popen(["open", str(p)])
        elif platform.system() == "Windows":
            os.startfile(str(p))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception:
        pass


def nome_cartella_risultati() -> str:
    # con i secondi: due calcoli ravvicinati non si sovrascrivono a vicenda
    return f"Risultati orario {datetime.now():%Y-%m-%d %H.%M.%S}"


# ── Finestra principale ──────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self.title("Orario pomeridiano - Liceo Musicale")
        self.geometry("1400x840")
        self.minsize(1000, 640)

        self.percorso: Path | None = None
        self.tabelle: dict[str, Tabella] = {}
        self.fogli: dict[str, Sheet] = {}
        self.orario: Orario | None = None
        self.coda: queue.Queue = queue.Queue()
        self._calcola_dopo_trasporti = False
        self._sto_sistemando = False
        self.modificato = False
        self._elenco_suggerito: tuple = ()

        self.ultima_cartella_risultati: Path | None = None
        imp = _carica_impostazioni()
        if imp.get("ultima_cartella_risultati") and Path(imp["ultima_cartella_risultati"]).exists():
            self.ultima_cartella_risultati = Path(imp["ultima_cartella_risultati"])

        self._crea_menu()
        self.schermata_apertura = self._crea_apertura()
        self.schermata_dati: ctk.CTkFrame | None = None
        self.schermata_apertura.pack(fill="both", expand=True)

    # ── Menu in alto ──
    def _crea_menu(self) -> None:
        barra = tk.Menu(self)
        m_file = tk.Menu(barra, tearoff=0)
        m_file.add_command(label="Nuovo file…", command=self.nuovo_file, accelerator="Cmd+N" if platform.system() == "Darwin" else "Ctrl+N")
        m_file.add_command(label="Apri file…", command=self.scegli_file, accelerator="Cmd+O" if platform.system() == "Darwin" else "Ctrl+O")
        m_file.add_command(label="Riapri ultimo file", command=self.riapri_ultimo)
        m_file.add_separator()
        m_file.add_command(label="Salva", command=self.salva_da_menu, accelerator="Cmd+S" if platform.system() == "Darwin" else "Ctrl+S")
        m_file.add_command(label="Salva con nome…", command=self.salva_con_nome)
        m_file.add_separator()
        m_file.add_command(label="Chiudi file", command=self.torna_apertura)
        m_file.add_command(label="Esci", command=self.destroy)
        barra.add_cascade(label="File", menu=m_file)

        m_mod = tk.Menu(barra, tearoff=0)
        m_mod.add_command(label="Aggiungi riga al foglio corrente", command=self.aggiungi_riga_corrente)
        m_mod.add_command(label="Elimina righe selezionate", command=self.elimina_righe_corrente)
        m_mod.add_separator()
        m_mod.add_command(label="Compila i docenti dagli studenti", command=self.compila_docenti)
        m_mod.add_command(label="Copia i nomi negli impegni studenti", command=self.compila_impegni)
        m_mod.add_separator()
        m_mod.add_command(label="Suggerimento: doppio clic su una cella per modificarla; tasto destro per altre azioni", state="disabled")
        barra.add_cascade(label="Modifica", menu=m_mod)

        m_or = tk.Menu(barra, tearoff=0)
        m_or.add_command(label="Calcola orario…", command=self.calcola_da_menu)
        m_or.add_command(label="Ricalcola da zero (ignora l'orario precedente)…", command=self.ricalcola_da_zero)
        m_or.add_separator()
        m_or.add_command(label="Aggiorna trasporti (tempi di ritorno a casa con i mezzi, serve internet)…",
                         command=self.aggiorna_trasporti)
        m_or.add_separator()
        m_or.add_command(label="Apri l'ultima cartella dei risultati", command=self.apri_risultati)
        barra.add_cascade(label="Orario", menu=m_or)

        m_aiuto = tk.Menu(barra, tearoff=0)
        m_aiuto.add_command(label="Istruzioni per compilare il file", command=self.mostra_istruzioni)
        m_aiuto.add_command(label="Informazioni", command=lambda: messagebox.showinfo(
            "Orario pomeridiano - Liceo Musicale",
            f"Versione {VERSIONE}\n\nCalcola l'orario pomeridiano del liceo musicale a partire da un file Excel "
            f"con studenti, docenti e gruppi, e produce Excel e PDF.\n\nMotore di calcolo: OR-Tools CP-SAT."))
        barra.add_cascade(label="Aiuto", menu=m_aiuto)
        self.config(menu=barra)
        mod = "Command" if platform.system() == "Darwin" else "Control"
        self.bind_all(f"<{mod}-n>", lambda e: self.nuovo_file())
        self.bind_all(f"<{mod}-o>", lambda e: self.scegli_file())
        self.bind_all(f"<{mod}-s>", lambda e: self.salva_da_menu())

    def _con_file(self) -> bool:
        if self.percorso is None or self.schermata_dati is None:
            messagebox.showinfo("Nessun file aperto", "Apri o crea prima un file dal menu File.")
            return False
        return True

    def salva_da_menu(self) -> None:
        if self._con_file():
            self.salva()

    def calcola_da_menu(self) -> None:
        if self._con_file():
            self.calcola()

    def ricalcola_da_zero(self) -> None:
        if self._con_file():
            self.usa_precedente.set(False)
            self.calcola()

    def aggiungi_riga_corrente(self) -> None:
        if self._con_file():
            nome = self.tabs.get()
            if nome in self.fogli:
                self.aggiungi_riga(nome)

    def elimina_righe_corrente(self) -> None:
        if self._con_file():
            nome = self.tabs.get()
            if nome in self.fogli:
                self.elimina_righe(nome)

    def riapri_ultimo(self) -> None:
        ultimo = _carica_impostazioni().get("ultimo_file")
        if ultimo and Path(ultimo).exists():
            self.apri_file(Path(ultimo))
        else:
            messagebox.showinfo("Nessun file recente", "Non c'è un file aperto di recente.")

    def mostra_istruzioni(self) -> None:
        FinestraTesto(self, "Istruzioni per compilare il file", "\n".join(ISTRUZIONI))

    def nuovo_file(self) -> None:
        iniziale = _carica_impostazioni().get("ultimo_file")
        p = filedialog.asksaveasfilename(
            title="Dove salvare il nuovo file di input",
            initialdir=str(Path(iniziale).parent) if iniziale else str(Path.home()),
            initialfile="orario_2026-27.xlsx", defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
        )
        if not p:
            return
        percorso = Path(p)
        try:
            crea_nuovo_file(percorso)
        except Exception as e:
            self.mostra_problemi([Problema("errore", percorso.name, f"Impossibile creare il file: {e}")], "Nuovo file")
            return
        self.apri_file(percorso)
        if self.schermata_dati is not None:
            self.lbl_stato.configure(text="Nuovo file creato: le righe grigie sono esempi da cancellare o sostituire.")

    def salva_con_nome(self) -> None:
        if not self._con_file():
            return
        assert self.percorso is not None
        p = filedialog.asksaveasfilename(
            title="Salva con nome", initialdir=str(self.percorso.parent), initialfile=self.percorso.name,
            defaultextension=".xlsx", filetypes=[("File Excel", "*.xlsx")],
        )
        if not p:
            return
        nuovo = Path(p)
        if nuovo != self.percorso:
            shutil.copy(self.percorso, nuovo)
            self.percorso = nuovo
            _salva_impostazioni({**_carica_impostazioni(), "ultimo_file": str(nuovo)})
            self.lbl_file.configure(text=f"📄  {nuovo.name}")
        self.salva()

    # ── Schermata 1: apertura ──
    def _crea_apertura(self) -> ctk.CTkFrame:
        f = ctk.CTkFrame(self, fg_color="transparent")
        box = ctk.CTkFrame(f, corner_radius=16)
        box.place(relx=0.5, rely=0.45, anchor="center")
        ctk.CTkLabel(box, text="Orario pomeridiano\nLiceo Musicale",
                     font=ctk.CTkFont(size=30, weight="bold"), justify="center").pack(padx=60, pady=(40, 10))
        ctk.CTkLabel(box, text="Seleziona il file Excel con studenti, docenti, gruppi e disponibilità.",
                     font=ctk.CTkFont(size=14), text_color="gray40").pack(padx=40, pady=(0, 24))
        ctk.CTkButton(box, text="Apri file Excel…", height=48, width=280,
                      font=ctk.CTkFont(size=16, weight="bold"), command=self.scegli_file).pack(pady=(0, 10))
        ctk.CTkButton(box, text="Nuovo file…", height=40, width=280, fg_color="#5a6b7d", hover_color="#48566a",
                      font=ctk.CTkFont(size=14), command=self.nuovo_file).pack(pady=(0, 14))
        ultimo = _carica_impostazioni().get("ultimo_file")
        if ultimo and Path(ultimo).exists():
            ctk.CTkButton(box, text=f"Riapri: {Path(ultimo).name}", height=36, width=280, fg_color="gray75",
                          text_color="black", hover_color="gray65",
                          command=lambda: self.apri_file(Path(ultimo))).pack(pady=(0, 30))
        else:
            ctk.CTkLabel(box, text="").pack(pady=(0, 20))
        return f

    def scegli_file(self) -> None:
        iniziale = _carica_impostazioni().get("ultimo_file")
        p = filedialog.askopenfilename(
            title="Scegli il file di input",
            initialdir=str(Path(iniziale).parent) if iniziale else str(Path.home()),
            filetypes=[("File Excel", "*.xlsx"), ("Tutti i file", "*.*")],
        )
        if p:
            self.apri_file(Path(p))

    def apri_file(self, percorso: Path) -> None:
        try:
            tabelle = leggi_tabelle(percorso)
        except ProblemiError as e:
            self.mostra_problemi(e.problemi, titolo="Il file non si può aprire")
            return
        except Exception as e:  # file corrotto, aperto in Excel, ecc.
            self.mostra_problemi([Problema("errore", percorso.name, f"Impossibile leggere il file: {e}")],
                                 titolo="Il file non si può aprire")
            return
        modifiche = aggiorna_struttura(tabelle)
        self.percorso = percorso
        self.tabelle = tabelle
        self.orario = None
        _salva_impostazioni({**_carica_impostazioni(), "ultimo_file": str(percorso)})
        self.schermata_apertura.pack_forget()
        if self.schermata_dati is not None:
            self.schermata_dati.destroy()
        self.schermata_dati = self._crea_dati()
        self.schermata_dati.pack(fill="both", expand=True)
        if modifiche:
            self.lbl_stato.configure(text=f"File aggiornato alla versione nuova: {len(modifiche)} aggiunte. "
                                          "Premi Salva per tenerle.")
            FinestraTesto(self, "File di una versione precedente",
                          "Il programma ha aggiornato la struttura del file senza toccare i dati:\n\n"
                          + "\n".join(f"   • {m}" for m in modifiche)
                          + "\n\nLe aggiunte diventano definitive quando premi Salva.",
                          sottotitolo=f"{len(modifiche)} aggiunte: fogli e colonne introdotti dopo che "
                                      "questo file era stato creato")

    # ── Schermata 2: dati ──
    def _crea_dati(self) -> ctk.CTkFrame:
        assert self.percorso is not None
        f = ctk.CTkFrame(self, fg_color="transparent")

        barra = ctk.CTkFrame(f, corner_radius=0, fg_color=("gray92", "gray18"), height=56)
        barra.pack(fill="x")
        self.lbl_file = ctk.CTkLabel(barra, text=f"📄  {self.percorso.name}", font=ctk.CTkFont(size=15, weight="bold"))
        self.lbl_file.pack(side="left", padx=16, pady=12)
        self.lbl_stato = ctk.CTkLabel(barra, text="", text_color="gray40")
        self.lbl_stato.pack(side="left", padx=8)

        ctk.CTkButton(barra, text="Cambia file", width=110, fg_color="gray70", text_color="black",
                      hover_color="gray60", command=self.torna_apertura).pack(side="right", padx=(4, 16), pady=10)
        self.btn_cartella = ctk.CTkButton(barra, text="Apri cartella risultati", width=170, fg_color="gray70",
                                          text_color="black", hover_color="gray60",
                                          state="normal" if self.ultima_cartella_risultati else "disabled",
                                          command=self.apri_risultati)
        self.btn_cartella.pack(side="right", padx=4, pady=10)
        self.btn_calcola = ctk.CTkButton(barra, text="⚡  Calcola orario", width=170, height=36,
                                         font=ctk.CTkFont(size=14, weight="bold"), fg_color="#1f7a3f",
                                         hover_color="#186331", command=self.calcola)
        self.btn_calcola.pack(side="right", padx=4, pady=10)
        self.btn_trasporti = ctk.CTkButton(barra, text="🚌  Trasporti", width=140, height=36, fg_color="#3a6ea5",
                                           hover_color="#2f5a86", command=self.aggiorna_trasporti)
        self.btn_trasporti.pack(side="right", padx=4, pady=10)
        ctk.CTkButton(barra, text="💾  Salva", width=110, height=36, command=self.salva).pack(side="right", padx=4, pady=10)

        self.usa_precedente = ctk.BooleanVar(value=True)
        self.chk_precedente = ctk.CTkCheckBox(barra, text="Parti dall'orario già calcolato (sposta solo il necessario)",
                                              variable=self.usa_precedente)
        self.chk_precedente.pack(side="right", padx=12)
        self._aggiorna_info_precedente()
        self.after(200, self._aggiorna_info_trasporti)
        self.after(250, self._aggiorna_suggerimenti)

        self.tabs = ctk.CTkTabview(f, anchor="nw", command=self._cambio_scheda)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.fogli = {}
        for nome in FOGLI_DATI:
            tab = self.tabs.add(nome)
            self._crea_foglio(tab, nome)
        return f

    def _stato_trasporti(self) -> tuple[int, int]:
        """(studenti con tempi dei mezzi aggiornati, studenti con indirizzo)."""
        if self.percorso is None:
            return 0, 0
        try:
            dati = costruisci_dati(self.tabelle, self.percorso)
        except Exception:
            return 0, 0
        con_indirizzo = [s for s in dati.studenti if s.indirizzo_completo]
        n = applica_trasporti(dati, leggi_trasporti(self.percorso))
        return n, len(con_indirizzo)

    def _aggiorna_info_trasporti(self) -> None:
        # può arrivare in ritardo (after) quando la schermata dati è già stata chiusa
        if not hasattr(self, "btn_trasporti") or not self.btn_trasporti.winfo_exists() or self.percorso is None:
            return
        pronti, totali = self._stato_trasporti()
        if totali == 0:
            self.btn_trasporti.configure(text="🚌  Trasporti", state="disabled")
        elif pronti >= totali:
            self.btn_trasporti.configure(text=f"🚌  Trasporti ✓ {pronti}", state="normal", fg_color="#4a7c59",
                                         hover_color="#3d6749")
        else:
            self.btn_trasporti.configure(text=f"🚌  Trasporti {pronti}/{totali}", state="normal", fg_color="#3a6ea5",
                                         hover_color="#2f5a86")

    def _cambio_scheda(self) -> None:
        """Aprendo la scheda dei gruppi si aggiorna l'elenco dei nomi suggeriti."""
        try:
            if self.tabs.get() in (FOGLIO_GRUPPI, FOGLIO_ABBINAMENTI, FOGLIO_LMI):
                self._aggiorna_suggerimenti()
        except Exception:
            pass

    def _aggiorna_info_precedente(self) -> None:
        prec = None
        try:
            prec = leggi_orario_precedente(self.percorso) if self.percorso else None
        except Exception:
            prec = None
        if prec:
            self.chk_precedente.configure(state="normal", text=f"Parti dall'orario calcolato il {prec[1]} (sposta solo il necessario)")
            self.usa_precedente.set(True)
        else:
            self.usa_precedente.set(False)
            self.chk_precedente.configure(state="disabled", text="Nessun orario precedente nel file")

    def _crea_foglio(self, parent, nome: str) -> None:
        t = self.tabelle[nome]
        comandi = ctk.CTkFrame(parent, fg_color="transparent")
        comandi.pack(fill="x", pady=(0, 4))
        ctk.CTkButton(comandi, text="+ Aggiungi riga", width=130, height=28,
                      command=lambda n=nome: self.aggiungi_riga(n)).pack(side="left", padx=(0, 6))
        ctk.CTkButton(comandi, text="− Elimina righe selezionate", width=190, height=28, fg_color="gray70",
                      text_color="black", hover_color="gray60",
                      command=lambda n=nome: self.elimina_righe(n)).pack(side="left")
        if nome == FOGLIO_DOCENTI:
            ctk.CTkButton(comandi, text="👥  Compila dagli studenti", width=200, height=28,
                          fg_color="#3a6ea5", hover_color="#2f5a86",
                          command=self.compila_docenti).pack(side="left", padx=(6, 0))
            ctk.CTkButton(comandi, text="Σ  Aggiorna i totali", width=160, height=28, fg_color="gray70",
                          text_color="black", hover_color="gray60",
                          command=self.aggiorna_totali).pack(side="left", padx=(6, 0))
        if nome == FOGLIO_IMPEGNI:
            ctk.CTkButton(comandi, text="👥  Copia i nomi dagli studenti", width=220, height=28,
                          fg_color="#3a6ea5", hover_color="#2f5a86",
                          command=self.compila_impegni).pack(side="left", padx=(6, 0))
        suggerimento = {
            FOGLIO_STUDENTI: "Le celle gialle vanno completate. Si incolla da Excel con "
                             + ("⌘V" if platform.system() == "Darwin" else "Ctrl+V") + ".",
            FOGLIO_DOCENTI: "X nelle ore disponibili, A nelle ore di accompagnamento fissate, vuoto se non disponibile.",
            FOGLIO_IMPEGNI: "Al contrario dei docenti: la X segna l'ora in cui il ragazzo NON può venire. "
                            "Chi non ha impegni si lascia vuoto.",
            FOGLIO_ABBINAMENTI: "Lezioni già decise che il programma deve rispettare. "
                                "Il tipo si può lasciare vuoto: lo deduce dal docente.",
        }.get(nome, "Si lavora come in Excel: copia e incolla, Invio scende, Tab va a destra, tasto destro per le righe.")
        ctk.CTkLabel(comandi, text=suggerimento, text_color="gray45").pack(side="left", padx=16)

        sheet = Sheet(parent, headers=list(t.intestazione), data=[list(r) for r in t.righe],
                      theme="light blue", show_row_index=True, header_font=(CARATTERE, 12, "bold"),
                      font=(CARATTERE, 12, "normal"))
        sheet.enable_bindings(*BINDINGS)
        sheet.set_options(**OPZIONI_FOGLIO)
        sheet.pack(fill="both", expand=True)
        self._adatta_colonne(sheet, nome)
        self.fogli[nome] = sheet
        # una riga vuota in fondo, sempre pronta: scrivendoci dentro ne compare un'altra
        sheet.bind("<<SheetModified>>", lambda _e, n=nome: self._foglio_modificato(n))
        sheet.extra_bindings("end_edit_cell", lambda ev, n=nome: self._cella_modificata(n, ev))
        self._riga_libera_in_fondo(nome)
        # ingrandire e rimpicciolire con ⌘/Ctrl + rotella o gesto del trackpad
        for combinazione in ("<Command-MouseWheel>", "<Control-MouseWheel>"):
            for widget in (sheet.MT, sheet.RI, sheet.CH):
                widget.bind(combinazione, lambda e, sh=sheet: self._zoom(sh, e.delta))

    def _zoom(self, sheet: Sheet, delta: float) -> str:
        try:
            sheet.zoom_in() if delta > 0 else sheet.zoom_out()
        except Exception:
            pass
        return "break"

    def _riga_libera_in_fondo(self, nome: str) -> None:
        """Tiene una riga vuota in coda, come il foglio infinito di Excel.

        Le righe vuote non vengono salvate: `tabelle_correnti` le scarta.
        """
        sheet = self.fogli.get(nome)
        if sheet is None or self._sto_sistemando:
            return
        if getattr(getattr(sheet.MT, "text_editor", None), "open", False):
            # si sta scrivendo in una cella: inserire una riga adesso chiuderebbe l'editor
            self.after(400, lambda: self._riga_libera_in_fondo(nome))
            return
        try:
            dati = sheet.get_sheet_data()
            n_col = len(self.tabelle[nome].intestazione)
            ultima_piena = not dati or any(str(v).strip() for v in dati[-1] if v is not None)
            if ultima_piena:
                self._sto_sistemando = True
                selezione = sheet.get_currently_selected()
                sheet.insert_row([""] * n_col, redraw=True)
                if selezione:  # inserire una riga non deve spostare il cursore
                    try:
                        sheet.select_cell(selezione.row, selezione.column, redraw=True)
                    except Exception:
                        pass
        except Exception:
            pass
        finally:
            self._sto_sistemando = False

    def _cella_modificata(self, nome: str, evento) -> None:
        """Fine modifica di una cella: segna il file come da salvare e completa il nome scritto."""
        self._foglio_modificato(nome)
        if nome not in (FOGLIO_GRUPPI, FOGLIO_LMI, FOGLIO_ABBINAMENTI):
            return
        r = getattr(evento, "row", None)
        c = getattr(evento, "column", None)
        if r is None or c is None:
            posizione = getattr(evento, "loc", None)
            if posizione is not None and len(posizione) >= 2:
                r, c = posizione[0], posizione[1]
        if r is not None and c is not None:
            self.after_idle(lambda: self._completa_nomi(nome, r, c))

    def _foglio_modificato(self, nome: str) -> None:
        self.modificato = True
        if hasattr(self, "lbl_stato"):
            self.lbl_stato.configure(text="Modifiche non salvate")
        self.after_idle(lambda: self._riga_libera_in_fondo(nome))

    def _adatta_colonne(self, sheet: Sheet, nome: str) -> None:
        try:
            sheet.set_all_cell_sizes_to_text()
            if nome == FOGLIO_DOCENTI:
                # le 20 colonne di disponibilità strette
                for i, h in enumerate(sheet.headers()):
                    if isinstance(h, str) and len(h) == 9 and h[3] == " " and ":" in h:
                        sheet.column_width(i, 48)
        except Exception:
            pass

    # ── Nomi degli studenti: suggerimento e completamento ────────────────────
    def _elenco_studenti(self) -> tuple[list[str], dict[str, list[str]]]:
        """(elenco "Cognome Nome" ordinato, cognome → nomi completi con quel cognome)."""
        try:
            ts = self.tabelle_correnti()[FOGLIO_STUDENTI]
        except Exception:
            return [], {}
        elenco: list[str] = []
        per_cognome: dict[str, list[str]] = {}
        for r in ts.righe:
            cognome = ts.valore(r, "Cognome").strip()
            nome = ts.valore(r, "Nome").strip()
            if not cognome:
                continue
            completo = f"{cognome.title()} {nome.title()}".strip()
            elenco.append(completo)
            per_cognome.setdefault(cognome.upper(), []).append(completo)
        return sorted(set(elenco)), per_cognome

    def _completa(self, testo: str, per_cognome: dict[str, list[str]]) -> str:
        """"Rossi" → "Rossi Mario" se il cognome è di uno solo; altrimenti lascia com'è."""
        t = " ".join(testo.split())
        if not t:
            return t
        omonimi = per_cognome.get(t.upper())
        if omonimi and len(omonimi) == 1:
            return omonimi[0]
        # cognome composto scritto per intero, senza nome
        for cognome, nomi in per_cognome.items():
            if len(nomi) == 1 and cognome == t.upper():
                return nomi[0]
        return t

    def _elenco_docenti(self) -> list[str]:
        try:
            td = self.tabelle_correnti()[FOGLIO_DOCENTI]
        except Exception:
            return []
        return sorted({td.valore(r, "Docente").strip() for r in td.righe if td.valore(r, "Docente").strip()})

    def _aggiorna_suggerimenti(self) -> None:
        """Menu a tendina sui fogli che nominano studenti, docenti, giorni e ore.

        Il menu è solo un suggerimento: `edit_data=False` perché altrimenti la libreria
        scriverebbe il primo valore dell'elenco in tutte le celle.
        """
        elenco, _ = self._elenco_studenti()
        docenti = self._elenco_docenti()
        if not elenco and not docenti:
            return
        chiave = (tuple(elenco), tuple(docenti))
        if chiave == self._elenco_suggerito:
            return
        # foglio → (inizio del nome della colonna, valori suggeriti)
        da_fare = {
            FOGLIO_GRUPPI: [("studente", elenco), ("docente", docenti)],
            FOGLIO_ABBINAMENTI: [("studente", elenco), ("docente", docenti),
                                 ("tipo", list(NOMI_TIPO)), ("giorno", list(GIORNI_LUNGHI)),
                                 ("ora", list(ORE))],
            FOGLIO_LMI: [("docente", docenti)],
        }
        try:
            for nome_foglio, regole in da_fare.items():
                sheet = self.fogli.get(nome_foglio)
                if sheet is None or not sheet.winfo_exists():
                    continue
                intest = self.tabelle[nome_foglio].intestazione
                for prefisso, valori in regole:
                    if not valori:
                        continue
                    for c, h in enumerate(intest):
                        if h.strip().lower().startswith(prefisso):
                            lettera = get_column_letter(c + 1)
                            sheet.dropdown(f"{lettera}:{lettera}", values=valori, state="normal",
                                           validate_input=False, edit_data=False, redraw=False)
                sheet.redraw()
            self._elenco_suggerito = chiave
        except Exception:
            pass

    def _completa_nomi(self, nome_foglio: str, riga: int | None = None, colonna: int | None = None) -> None:
        """Espande in "Cognome Nome" il cognome appena scritto, se non ci sono omonimi.

        Tocca solo la cella modificata: così non si perde la selezione e la griglia non lampeggia.
        """
        sheet = self.fogli.get(nome_foglio)
        if sheet is None or self._sto_sistemando or riga is None or colonna is None:
            return
        intest = self.tabelle[nome_foglio].intestazione
        if colonna >= len(intest):
            return
        titolo = intest[colonna].strip().lower()
        if nome_foglio == FOGLIO_LMI:
            if not titolo.startswith("studenti"):
                return
        elif not titolo.startswith("studente"):
            return
        _, per_cognome = self._elenco_studenti()
        if not per_cognome:
            return
        try:
            testo = str(sheet.get_cell_data(riga, colonna) or "")
            if not testo.strip():
                return
            if nome_foglio == FOGLIO_LMI:
                pezzi = [self._completa(x, per_cognome) for x in testo.split(",")]
                nuovo = ", ".join(x for x in pezzi if x)
            else:
                nuovo = self._completa(testo, per_cognome)
            if nuovo != testo:
                self._sto_sistemando = True
                sheet.set_cell_data(riga, colonna, nuovo, redraw=True)
        except Exception:
            pass
        finally:
            self._sto_sistemando = False

    def aggiorna_totali(self) -> None:
        """Ricalcola la colonna «Ore dichiarate» e la riga TOTALE del foglio Docenti."""
        if not self._con_file():
            return
        tabelle = self.tabelle_correnti()
        aggiorna_totali_docenti(tabelle)
        sheet = self.fogli[FOGLIO_DOCENTI]
        sheet.set_sheet_data([list(r) for r in tabelle[FOGLIO_DOCENTI].righe], redraw=True)
        self._adatta_colonne(sheet, FOGLIO_DOCENTI)
        self._riga_libera_in_fondo(FOGLIO_DOCENTI)
        self.modificato = True
        td = tabelle[FOGLIO_DOCENTI]
        totale = td.righe[-1][td.colonna("Ore dichiarate")] if td.righe else "0"
        self.lbl_stato.configure(text=f"Totali aggiornati: {totale} ore dichiarate in tutto")

    def compila_docenti(self) -> None:
        """Riempie il foglio Docenti con i nomi che compaiono negli studenti e nei gruppi.

        Aggiunge i docenti mancanti e completa la colonna degli strumenti. Non tocca aule,
        disponibilità, ore di accompagnamento e note già inseriti.
        """
        if not self._con_file():
            return
        tabelle = self.tabelle_correnti()
        ts, tg, td = tabelle[FOGLIO_STUDENTI], tabelle.get(FOGLIO_GRUPPI), tabelle[FOGLIO_DOCENTI]

        # docente → strumenti che insegna
        trovati: dict[str, set[str]] = {}
        for r in ts.righe:
            for col_doc, col_strum in (("Docente 1", "Strumento 1"), ("Docente 2", "Strumento 2")):
                nome = ts.valore(r, col_doc)
                if nome:
                    trovati.setdefault(nome, set()).add(ts.valore(r, col_strum).upper())
        if tg is not None:
            for r in tg.righe:
                nome = tg.valore(r, "Docente")
                if nome and any(tg.valore(r, f"Studente {k}") for k in range(1, 6)):
                    trovati.setdefault(nome, set()).add("MUSICA DA CAMERA")
        if not trovati:
            messagebox.showinfo("Compila dagli studenti",
                                "Nel foglio Studenti non c'è nessun docente: compila prima le colonne "
                                "Docente 1 e Docente 2.")
            return

        i_nome = td.colonna("Docente")
        i_strum = td.colonna("Strumento")
        n_col = len(td.intestazione)
        gia_presenti = {r[i_nome].strip(): r for r in td.righe if i_nome < len(r) and r[i_nome].strip()}

        aggiunti, aggiornati = [], []
        for nome in sorted(trovati):
            strumenti = sorted(x for x in trovati[nome] if x)
            riga = gia_presenti.get(nome)
            if riga is None:
                nuova = [""] * n_col
                nuova[i_nome] = nome
                nuova[i_strum] = ", ".join(strumenti)
                td.righe.append(nuova)
                aggiunti.append(nome)
            else:
                vecchi = {x.strip().upper() for x in (riga[i_strum] if i_strum < len(riga) else "").split(",") if x.strip()}
                unione = sorted(vecchi | set(strumenti))
                if unione != sorted(vecchi):
                    riga[i_strum] = ", ".join(unione)
                    aggiornati.append(nome)
        non_usati = [n for n in gia_presenti if n not in trovati]

        td.righe.sort(key=lambda r: r[i_nome].strip().lower() if i_nome < len(r) else "")
        self.fogli[FOGLIO_DOCENTI].set_sheet_data([list(r) for r in td.righe], redraw=True)
        self._adatta_colonne(self.fogli[FOGLIO_DOCENTI], FOGLIO_DOCENTI)
        self._riga_libera_in_fondo(FOGLIO_DOCENTI)
        self.tabs.set(FOGLIO_DOCENTI)
        self.modificato = True

        parti = [f"Docenti trovati negli studenti e nei gruppi: {len(trovati)}."]
        parti.append(f"Aggiunti {len(aggiunti)}: " + ", ".join(aggiunti) if aggiunti else "Nessun docente da aggiungere.")
        if aggiornati:
            parti.append(f"Strumenti completati per {len(aggiornati)}: " + ", ".join(aggiornati) + ".")
        if non_usati:
            parti.append(f"Nel foglio Docenti ce ne sono {len(non_usati)} che nessuno studente ha come docente "
                         f"(nomi scritti diversamente?): " + ", ".join(non_usati) + ".")
        parti.append("\nRestano da compilare a mano le aule e le disponibilità. Premi Salva per tenere le modifiche.")
        messagebox.showinfo("Compila dagli studenti", "\n\n".join(parti))
        self.lbl_stato.configure(text=f"Docenti compilati: {len(aggiunti)} aggiunti, {len(aggiornati)} completati")

    def compila_impegni(self) -> None:
        """Ricopia l'elenco dei ragazzi dal foglio Studenti, lasciando le ore già segnate."""
        if not self._con_file():
            return
        tabelle = self.tabelle_correnti()
        ts, ti = tabelle[FOGLIO_STUDENTI], tabelle[FOGLIO_IMPEGNI]
        i_cl, i_cog, i_nome = ti.colonna("Classe"), ti.colonna("Cognome"), ti.colonna("Nome")
        n_col = len(ti.intestazione)
        gia = {(r[i_cog].strip().upper(), r[i_nome].strip().upper()) for r in ti.righe
               if i_cog < len(r) and r[i_cog].strip()}
        aggiunti = []
        for r in ts.righe:
            cognome, nome = ts.valore(r, "Cognome"), ts.valore(r, "Nome")
            if not cognome or (cognome.upper(), nome.upper()) in gia:
                continue
            nuova = [""] * n_col
            nuova[i_cl] = ts.valore(r, "Classe")
            nuova[i_cog], nuova[i_nome] = cognome, nome
            ti.righe.append(nuova)
            aggiunti.append(f"{cognome.title()} {nome.title()}")
        nel_foglio = {(ts.valore(r, "Cognome").upper(), ts.valore(r, "Nome").upper()) for r in ts.righe}
        estranei = [f"{r[i_cog]} {r[i_nome]}" for r in ti.righe
                    if i_cog < len(r) and r[i_cog].strip()
                    and (r[i_cog].strip().upper(), r[i_nome].strip().upper()) not in nel_foglio]

        def chiave(r):
            classe = r[i_cl] if i_cl < len(r) else ""
            return (int(classe) if str(classe).strip().isdigit() else 9,
                    r[i_cog].strip().lower() if i_cog < len(r) else "",
                    r[i_nome].strip().lower() if i_nome < len(r) else "")

        ti.righe.sort(key=chiave)
        self.fogli[FOGLIO_IMPEGNI].set_sheet_data([list(r) for r in ti.righe], redraw=True)
        self._adatta_colonne(self.fogli[FOGLIO_IMPEGNI], FOGLIO_IMPEGNI)
        self._riga_libera_in_fondo(FOGLIO_IMPEGNI)
        self.tabs.set(FOGLIO_IMPEGNI)
        self.modificato = True
        parti = [f"Ragazzi aggiunti all'elenco: {len(aggiunti)}." if aggiunti
                 else "L'elenco era già completo: nessun ragazzo da aggiungere."]
        if estranei:
            parti.append(f"Ci sono {len(estranei)} righe con nomi che non stanno nel foglio Studenti: "
                         + ", ".join(estranei[:8]) + (" …" if len(estranei) > 8 else "") + ".")
        parti.append("Ricorda: qui la X segna l'ora in cui il ragazzo NON può venire.")
        messagebox.showinfo("Copia i nomi dagli studenti", "\n\n".join(parti))
        self.lbl_stato.configure(text=f"Elenco impegni aggiornato: {len(aggiunti)} ragazzi aggiunti")

    def aggiungi_riga(self, nome: str) -> None:
        sheet = self.fogli[nome]
        n_col = len(self.tabelle[nome].intestazione)
        sheet.insert_row([""] * n_col, redraw=True)

    def elimina_righe(self, nome: str) -> None:
        sheet = self.fogli[nome]
        righe = sorted(sheet.get_selected_rows(get_cells_as_rows=True), reverse=True)
        if not righe:
            self.lbl_stato.configure(text="Seleziona prima una o più righe (clic sul numero a sinistra).")
            return
        sheet.del_rows(righe, redraw=True)

    def torna_apertura(self) -> None:
        # chiudere il file significa dimenticarlo: altrimenti Salva/Calcola dal menu
        # lavorerebbero su griglie non più visibili
        if self.schermata_dati is not None:
            self.schermata_dati.destroy()
            self.schermata_dati = None
        self.percorso = None
        self.tabelle = {}
        self.fogli = {}
        self.orario = None
        self.schermata_apertura.destroy()
        self.schermata_apertura = self._crea_apertura()
        self.schermata_apertura.pack(fill="both", expand=True)

    # ── Dati dalla griglia ──
    def tabelle_correnti(self) -> dict[str, Tabella]:
        out: dict[str, Tabella] = {}
        for nome, sheet in self.fogli.items():
            intest = list(self.tabelle[nome].intestazione)
            righe = []
            for r in sheet.get_sheet_data():
                vals = ["" if v is None else str(v).strip() for v in r][:len(intest)]
                vals += [""] * (len(intest) - len(vals))
                if any(vals):
                    righe.append(vals)
            out[nome] = Tabella(nome, intest, righe)
        return out

    def salva(self) -> bool:
        assert self.percorso is not None
        try:
            self.tabelle = self.tabelle_correnti()
            aggiorna_totali_docenti(self.tabelle)
            salva_tabelle(self.percorso, self.tabelle)
            sheet_doc = self.fogli.get(FOGLIO_DOCENTI)
            if sheet_doc is not None and sheet_doc.winfo_exists():
                sheet_doc.set_sheet_data([list(r) for r in self.tabelle[FOGLIO_DOCENTI].righe], redraw=True)
        except PermissionError:
            self.mostra_problemi([Problema("errore", self.percorso.name,
                                           "Il file è aperto in Excel: chiuderlo e salvare di nuovo.")], "Salvataggio non riuscito")
            return False
        except Exception as e:
            self.mostra_problemi([Problema("errore", self.percorso.name, f"Salvataggio non riuscito: {e}")],
                                 "Salvataggio non riuscito")
            return False
        self.modificato = False
        self.lbl_stato.configure(text=f"Salvato alle {datetime.now():%H:%M}")
        return True

    # ── Calcolo ──
    def calcola(self) -> None:
        assert self.percorso is not None
        if not self.salva():
            return
        # controlli subito: gli errori fermano qui, le approssimazioni si mostrano e si conferma
        try:
            from .controlli import controlla
            dati_controllati = costruisci_dati(self.tabelle, self.percorso)
            applica_trasporti(dati_controllati, leggi_trasporti(self.percorso))
            avvisi_prima = controlla(dati_controllati)
        except ProblemiError as e:
            self.mostra_problemi(e.problemi, titolo="Il calcolo si è fermato: ci sono problemi da risolvere")
            return
        except Exception:
            avvisi_prima = []
        approssimazioni = [a for a in avvisi_prima if a.categoria]
        if approssimazioni and not self._calcola_dopo_trasporti:
            if not FinestraApprossimazioni(self, approssimazioni).procedi:
                return

        # se ci sono indirizzi ma i tempi dei mezzi mancano, proponi di calcolarli prima
        if not self._calcola_dopo_trasporti:
            pronti, totali = self._stato_trasporti()
            if totali and pronti < totali:
                mancanti = totali - pronti
                risposta = messagebox.askyesnocancel(
                    "Tempi dei mezzi mancanti",
                    f"{mancanti} studenti su {totali} hanno l'indirizzo ma non i tempi di ritorno a casa con i mezzi.\n\n"
                    f"Sì = calcolo prima i trasporti (circa {max(1, round(mancanti * 3.2 / 60))} min, serve internet), "
                    f"poi l'orario\nNo = calcolo l'orario adesso usando i KM")
                if risposta is None:
                    return
                if risposta:
                    self._calcola_dopo_trasporti = True
                    self.aggiorna_trasporti(solo_nuovi_forzato=True)
                    return
        self._calcola_dopo_trasporti = False
        base = self.ultima_cartella_risultati.parent if self.ultima_cartella_risultati else self.percorso.parent
        scelta = filedialog.askdirectory(title="Dove salvare i risultati? (verrà creata una cartella con la data)",
                                         initialdir=str(base), mustexist=True)
        if not scelta:
            return
        cartella = Path(scelta) / nome_cartella_risultati()
        tabelle = self.tabelle
        percorso = self.percorso
        usa_prec = bool(self.usa_precedente.get())
        self.btn_calcola.configure(state="disabled")
        self.attesa = FinestraAttesa(self)

        def lavoro() -> None:
            try:
                from .controlli import controlla
                from .export import esporta_tutto
                from .motore import calcola as calcola_orario

                self.coda.put(("msg", "Lettura dei dati…"))
                dati = costruisci_dati(tabelle, percorso)
                applica_trasporti(dati, leggi_trasporti(percorso))
                self.coda.put(("msg", "Controllo dei dati…"))
                avvisi = controlla(dati)
                precedente = None
                if usa_prec:
                    letto = leggi_orario_precedente(percorso)
                    precedente = letto[0] if letto else None
                orario = calcola_orario(dati, progresso=lambda m: self.coda.put(("msg", m)), precedente=precedente)
                orario.avvisi = avvisi + orario.avvisi
                self.coda.put(("msg", "Scrittura di Excel e PDF…"))
                file_creati = esporta_tutto(orario, cartella)
                self.coda.put(("msg", "Salvataggio dell'orario nel file di input…"))
                salva_orario_nel_file(percorso, orario)
                self.coda.put(("ok", orario, file_creati, cartella))
            except ProblemiError as e:
                self.coda.put(("problemi", e.problemi))
            except Exception:
                self.coda.put(("crash", traceback.format_exc()))

        threading.Thread(target=lavoro, daemon=True).start()
        self.after(150, self._controlla_coda)

    def _controlla_coda(self) -> None:
        try:
            while True:
                ev = self.coda.get_nowait()
                if ev[0] == "msg":
                    self.attesa.messaggio(ev[1])
                    continue
                self.attesa.chiudi()
                # importando dalla schermata iniziale il pulsante Calcola non esiste ancora
                if hasattr(self, "btn_calcola") and self.btn_calcola.winfo_exists():
                    self.btn_calcola.configure(state="normal")
                if ev[0] == "ok":
                    self.orario = ev[1]
                    self.ultima_cartella_risultati = ev[3]
                    _salva_impostazioni({**_carica_impostazioni(), "ultima_cartella_risultati": str(ev[3])})
                    self.btn_cartella.configure(state="normal")
                    self._aggiorna_info_precedente()
                    self.mostra_orario(ev[1], ev[2], ev[3])
                elif ev[0] == "trasporti":
                    self._aggiorna_info_trasporti()
                    self.lbl_stato.configure(text="Trasporti aggiornati e salvati nel foglio 'Trasporti'.")
                    if self._calcola_dopo_trasporti and not ev[2]:
                        self.lbl_stato.configure(text=ev[1].split(chr(10))[0] + " Ora calcolo l'orario…")
                        self.after(300, self.calcola)
                        return
                    self._calcola_dopo_trasporti = False
                    testo = ev[1] + ("\n\n(Aggiornamento interrotto: i risultati parziali sono stati salvati.)" if ev[2] else "")
                    messagebox.showinfo("Trasporti aggiornati", testo)
                elif ev[0] == "problemi":
                    self._calcola_dopo_trasporti = False
                    self.mostra_problemi(ev[1], titolo="Il calcolo si è fermato: ci sono problemi da risolvere")
                else:
                    self._calcola_dopo_trasporti = False
                    self.mostra_problemi([Problema("errore", "Programma", "Errore inatteso:\n" + ev[1])],
                                         titolo="Errore inatteso")
                return
        except queue.Empty:
            pass
        self.after(150, self._controlla_coda)

    def aggiorna_trasporti(self, solo_nuovi_forzato: bool = False) -> None:
        if not self._con_file() or not self.salva():
            self._calcola_dopo_trasporti = False
            return
        assert self.percorso is not None
        percorso = self.percorso
        try:
            dati = costruisci_dati(self.tabelle, percorso)
        except ProblemiError as e:
            self.mostra_problemi(e.problemi, "Correggi prima questi errori nei dati")
            return
        con_indirizzo = [s for s in dati.studenti if s.indirizzo_completo]
        if not dati.parametri.indirizzo_scuola.strip():
            self._calcola_dopo_trasporti = False
            messagebox.showwarning("Manca l'indirizzo della scuola",
                                   "Nel foglio Parametri, alla riga 'Indirizzo della scuola', scrivi via, numero e comune "
                                   "della scuola. È il punto di partenza per calcolare i mezzi.")
            return
        if not con_indirizzo:
            self._calcola_dopo_trasporti = False
            messagebox.showwarning("Nessun indirizzo", "Nessuno studente ha Comune e Indirizzo compilati nel foglio Studenti.")
            return
        precedenti_ok = leggi_trasporti(percorso)
        gia = [s for s in con_indirizzo if (t := precedenti_ok.get(s.id)) and t.esito == "OK"
               and t.indirizzo_usato == s.indirizzo_completo]
        nuovi = len(con_indirizzo) - len(gia)
        solo_nuovi = False
        if solo_nuovi_forzato:
            solo_nuovi = True  # richiesto dal calcolo: rifà solo i mancanti, senza chiedere
        elif gia:
            risposta = messagebox.askyesnocancel(
                "Aggiorna trasporti",
                f"{len(gia)} studenti hanno già i tempi dei mezzi calcolati con lo stesso indirizzo; "
                f"{nuovi} sono nuovi o con indirizzo cambiato.\n\n"
                f"Sì = ricalcolo TUTTI ({max(1, round(len(con_indirizzo) * 3.2 / 60))} min circa, utile se gli orari dei "
                f"mezzi sono cambiati)\nNo = solo i {nuovi} nuovi o modificati ({max(1, round(nuovi * 3.2 / 60))} min circa)")
            if risposta is None:
                return
            solo_nuovi = not risposta
        elif not solo_nuovi_forzato and not messagebox.askyesno(
            "Aggiorna trasporti",
            f"Verranno cercati, tramite internet (OpenStreetMap e Transitous), i tempi di ritorno a casa con i mezzi "
            f"pubblici per {len(con_indirizzo)} studenti, per ciascuna delle 4 fasce.\n\n"
            f"Ci vogliono circa {max(1, round(len(con_indirizzo) * 3.2 / 60))} minuti. I risultati vengono salvati nel "
            f"foglio 'Trasporti' del file e usati dal calcolo al posto dei KM.\n\nProcedere?"):
            self._calcola_dopo_trasporti = False
            return
        self.btn_calcola.configure(state="disabled")
        annulla = threading.Event()
        self.attesa = FinestraAttesa(self, titolo="Aggiornamento trasporti", annulla=annulla)

        def lavoro() -> None:
            try:
                from .trasporti import TrasportiError, aggiorna_trasporti, riepilogo
                precedenti = leggi_trasporti(percorso)
                risultati = aggiorna_trasporti(dati, progresso=lambda m: self.coda.put(("msg", m)),
                                               annulla=annulla, precedenti=precedenti, solo_nuovi=solo_nuovi)
                # gli studenti non rifatti (interruzione) mantengono i dati vecchi
                uniti = {**precedenti, **risultati}
                salva_trasporti(percorso, uniti)
                self.coda.put(("trasporti", riepilogo(risultati, dati), annulla.is_set()))
            except TrasportiError as e:
                self.coda.put(("problemi", [Problema("errore", "Trasporti", str(e))]))
            except Exception:
                self.coda.put(("crash", traceback.format_exc()))

        threading.Thread(target=lavoro, daemon=True).start()
        self.after(150, self._controlla_coda)

    def apri_risultati(self) -> None:
        if self.ultima_cartella_risultati and self.ultima_cartella_risultati.exists():
            apri_cartella(self.ultima_cartella_risultati)
        else:
            messagebox.showinfo("Nessun risultato", "Non è ancora stato calcolato nessun orario.")

    # ── Risultato ──
    def mostra_orario(self, orario: Orario, file_creati: list[Path], cartella: Path) -> None:
        nome_tab = "Orario"
        try:
            self.tabs.delete(nome_tab)
        except Exception:
            pass
        tab = self.tabs.add(nome_tab)
        self.tabs.set(nome_tab)

        n_stud = len(orario.dati.studenti)
        rientri = [orario.rientri_di(s.id) for s in orario.dati.studenti]
        buchi = sum(orario.buchi_di(s.id) for s in orario.dati.studenti)
        riepilogo = (
            f"Orario calcolato ({'ottimo' if orario.stato == 'OPTIMAL' else 'buona soluzione'} in {orario.secondi:.0f} s).  "
            f"Studenti: {n_stud}.  Lezioni: {len(orario.lezioni)}.  Buchi totali: {buchi}.  "
            f"Rientri: {rientri.count(1)} studenti con 1, {rientri.count(2)} con 2, {rientri.count(3)} con 3.  "
            f"Avvisi: {len(orario.avvisi)}."
        )
        ctk.CTkLabel(tab, text=riepilogo, anchor="w", font=ctk.CTkFont(size=13)).pack(fill="x", padx=4, pady=(0, 4))

        sotto = ctk.CTkTabview(tab, anchor="nw")
        sotto.pack(fill="both", expand=True)
        docenti = orario.dati.docenti
        mappa = orario.per_docente_fascia()
        studenti = {s.id: s for s in orario.dati.studenti}
        cognomi_doppi = {s.cognome for s in orario.dati.studenti
                         if sum(1 for x in orario.dati.studenti if x.cognome == s.cognome) > 1}

        def testo(l) -> str:
            if l.tipo == TIPO_ACCOMP:
                return ETICHETTA_ACCOMP
            if l.tipo == TIPO_LMC:
                return "LMC: " + ", ".join(studenti[i].etichetta(studenti[i].cognome in cognomi_doppi) for i in l.studenti)
            s = studenti[l.studenti[0]]
            pre = "1° " if l.tipo == TIPO_STRUM1 else "2° "
            return pre + s.etichetta(s.cognome in cognomi_doppi) + (" (2ª ora)" if l.due_ore else "")

        for g, giorno in enumerate(GIORNI_LUNGHI):
            t = sotto.add(giorno)
            intest = [f"{d.nome}\n{d.aula(g)}" for d in docenti]
            righe = []
            for o in range(N_ORE):
                riga = []
                for d in docenti:
                    l = mappa.get((d.nome, fascia(g, o)))
                    riga.append(testo(l) if l else ("" if d.disponibile(fascia(g, o)) else "—"))
                righe.append(riga)
            sh = Sheet(t, headers=intest, data=righe, theme="light blue", show_row_index=True,
                       row_index=[f"{ORE_LABEL[o]} {ORE[o]}-{ORE_FINE[o]}" for o in range(N_ORE)],
                       font=(CARATTERE, 11, "normal"), header_font=(CARATTERE, 11, "bold"))
            sh.enable_bindings("single_select", "copy", "column_width_resize", "arrowkeys")
            sh.pack(fill="both", expand=True)
            try:
                sh.set_all_cell_sizes_to_text()
            except Exception:
                pass

        t = sotto.add("Avvisi")
        box = ctk.CTkTextbox(t, font=ctk.CTkFont(size=13))
        box.pack(fill="both", expand=True)
        if orario.avvisi:
            box.insert("end", "Preferenze non soddisfatte e segnalazioni (l'orario è comunque valido):\n\n")
            for a in orario.avvisi:
                box.insert("end", f"• {a.dove}: {a.messaggio}\n")
        else:
            box.insert("end", "Nessun avviso: tutte le preferenze sono soddisfatte.")
        box.configure(state="disabled")

        FinestraFatto(self, file_creati, cartella)

    def mostra_problemi(self, problemi: list[Problema], titolo: str) -> None:
        FinestraProblemi(self, problemi, titolo)


# ── Finestre secondarie ──────────────────────────────────────────────────────

class FinestraAttesa(ctk.CTkToplevel):
    def __init__(self, master, titolo: str = "Calcolo in corso", annulla: threading.Event | None = None) -> None:
        super().__init__(master)
        self.title(titolo)
        self.geometry("520x190" if annulla else "460x150")
        self.resizable(False, False)
        self.transient(master)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        self.lbl = ctk.CTkLabel(self, text="Avvio…", font=ctk.CTkFont(size=14), wraplength=480)
        self.lbl.pack(pady=(28, 12))
        self.bar = ctk.CTkProgressBar(self, mode="indeterminate", width=380)
        self.bar.pack()
        self.bar.start()
        if annulla is not None:
            ctk.CTkButton(self, text="Interrompi", fg_color="gray70", text_color="black", hover_color="gray60",
                          command=lambda: (annulla.set(), self.lbl.configure(text="Interruzione in corso…"))).pack(pady=14)
        self.after(50, self.grab_set)

    def messaggio(self, m: str) -> None:
        self.lbl.configure(text=m)

    def chiudi(self) -> None:
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


class FinestraProblemi(ctk.CTkToplevel):
    def __init__(self, master, problemi: list[Problema], titolo: str) -> None:
        super().__init__(master)
        self.title(titolo)
        self.geometry("900x560")
        self.transient(master)
        errori = [p for p in problemi if p.livello == "errore"]
        avvisi = [p for p in problemi if p.livello == "avviso"]
        ctk.CTkLabel(self, text=titolo, font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(16, 4), padx=16, anchor="w")
        ctk.CTkLabel(self, text=f"{len(errori)} errori da correggere, {len(avvisi)} avvisi. "
                          "Correggi i dati nelle griglie e premi di nuovo Calcola.",
                     text_color="gray40").pack(padx=16, anchor="w")
        box = ctk.CTkTextbox(self, font=ctk.CTkFont(size=13), wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=12)
        self.testo = ""
        if errori:
            self.testo += "ERRORI (bloccano il calcolo)\n\n"
            for p in errori:
                self.testo += f"✖ {p.dove}\n   {p.messaggio}\n\n"
        if avvisi:
            self.testo += "AVVISI (non bloccano)\n\n"
            for p in avvisi:
                self.testo += f"• {p.dove}\n   {p.messaggio}\n\n"
        box.insert("end", self.testo)
        box.configure(state="disabled")
        pulsanti = ctk.CTkFrame(self, fg_color="transparent")
        pulsanti.pack(pady=(0, 14))
        ctk.CTkButton(pulsanti, text="Copia il testo", fg_color="gray70", text_color="black", hover_color="gray60",
                      command=self.copia).pack(side="left", padx=6)
        ctk.CTkButton(pulsanti, text="Chiudi", command=self.destroy).pack(side="left", padx=6)
        self.after(50, self.grab_set)

    def copia(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.testo)


class FinestraTesto(ctk.CTkToplevel):
    def __init__(self, master, titolo: str, testo: str, sottotitolo: str = "") -> None:
        super().__init__(master)
        self.title(titolo)
        self.geometry("900x600")
        self.transient(master)
        if sottotitolo:
            ctk.CTkLabel(self, text=sottotitolo, font=ctk.CTkFont(size=14, weight="bold"),
                         wraplength=850, justify="left").pack(padx=16, pady=(16, 0), anchor="w")
        box = ctk.CTkTextbox(self, font=ctk.CTkFont(size=13), wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=16)
        box.insert("end", testo)
        box.configure(state="disabled")
        ctk.CTkButton(self, text="Chiudi", command=self.destroy).pack(pady=(0, 14))


class FinestraApprossimazioni(ctk.CTkToplevel):
    """Elenca, prima di calcolare, tutto ciò che il programma darà per buono."""

    def __init__(self, master, avvisi: list[Problema]) -> None:
        super().__init__(master)
        self.procedi = False
        self.title("Prima di calcolare")
        self.geometry("900x620")
        self.transient(master)

        gruppi: dict[str, list[Problema]] = {}
        for a in avvisi:
            gruppi.setdefault(a.categoria, []).append(a)
        ordinati = sorted(gruppi.items(), key=lambda kv: -len(kv[1]))

        ctk.CTkLabel(self, text="Il programma calcolerà lo stesso, dando per buone queste cose",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(16, 2), padx=16, anchor="w")
        ctk.CTkLabel(self, text="Nessuna di queste impedisce il calcolo. Se qualcuna non ti convince, "
                                "annulla, correggi i dati e ricalcola.",
                     text_color="gray40", wraplength=850, justify="left").pack(padx=16, anchor="w")

        riassunto = ctk.CTkFrame(self, fg_color=("gray92", "gray18"))
        riassunto.pack(fill="x", padx=16, pady=10)
        for categoria, elenco in ordinati:
            ctk.CTkLabel(riassunto, text=f"•  {len(elenco)}  {categoria}",
                         font=ctk.CTkFont(size=14), anchor="w").pack(fill="x", padx=12, pady=3)

        box = ctk.CTkTextbox(self, font=ctk.CTkFont(size=12), wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        self.testo = ""
        for categoria, elenco in ordinati:
            self.testo += f"{categoria.upper()} ({len(elenco)})\n"
            for a in elenco:
                self.testo += f"   • {a.dove}: {a.messaggio}\n"
            self.testo += "\n"
        box.insert("end", self.testo)
        box.configure(state="disabled")

        pulsanti = ctk.CTkFrame(self, fg_color="transparent")
        pulsanti.pack(pady=(0, 14))
        ctk.CTkButton(pulsanti, text="Annulla, correggo i dati", width=200, fg_color="gray70",
                      text_color="black", hover_color="gray60", command=self.destroy).pack(side="left", padx=6)
        ctk.CTkButton(pulsanti, text="Copia il testo", width=140, fg_color="gray70", text_color="black",
                      hover_color="gray60", command=self.copia).pack(side="left", padx=6)
        ctk.CTkButton(pulsanti, text="⚡  Calcola comunque", width=200, height=36, fg_color="#1f7a3f",
                      hover_color="#186331", font=ctk.CTkFont(size=14, weight="bold"),
                      command=self.avanti).pack(side="left", padx=6)
        self.after(50, self.grab_set)
        master.wait_window(self)

    def copia(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.testo)

    def avanti(self) -> None:
        self.procedi = True
        self.destroy()


class FinestraFatto(ctk.CTkToplevel):
    def __init__(self, master, file_creati: list[Path], cartella: Path) -> None:
        super().__init__(master)
        self.title("Orario creato")
        self.geometry("560x300")
        self.transient(master)
        ctk.CTkLabel(self, text="✅  Orario creato", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(22, 6))
        ctk.CTkLabel(self, text=f"I file sono nella cartella:\n{cartella}", justify="center",
                     text_color="gray30").pack(pady=(0, 8))
        ctk.CTkLabel(self, text="\n".join(f.name for f in file_creati), justify="left",
                     font=ctk.CTkFont(size=13)).pack()
        pulsanti = ctk.CTkFrame(self, fg_color="transparent")
        pulsanti.pack(pady=18)
        ctk.CTkButton(pulsanti, text="Apri la cartella", command=lambda: apri_cartella(cartella)).pack(side="left", padx=6)
        ctk.CTkButton(pulsanti, text="OK", fg_color="gray70", text_color="black", hover_color="gray60",
                      command=self.destroy).pack(side="left", padx=6)


def avvia() -> None:
    app = App()
    app.mainloop()
