"""Finestra del programma: apri file → fogli a griglia modificabili → Calcola → Excel + PDF."""

from __future__ import annotations

import json
import os
import platform
import queue
import shutil
import subprocess
import threading
import tkinter as tk
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tksheet import Sheet

from .costanti import (
    ETICHETTA_ACCOMP, FOGLI_DATI, FOGLIO_DOCENTI, FOGLIO_STUDENTI, GIORNI_LUNGHI,
    N_ORE, ORE, ORE_FINE, ORE_LABEL, TIPO_ACCOMP, TIPO_LMC, TIPO_STRUM1, fascia,
)
from .lettura import (
    Tabella, aggiorna_struttura, applica_trasporti, costruisci_dati, leggi_orario_precedente, leggi_tabelle, leggi_trasporti,
    salva_orario_nel_file, salva_tabelle, salva_trasporti,
)
from .modello import Orario, Problema, ProblemiError
from .template import ISTRUZIONI, crea_nuovo_file

VERSIONE = "1.0 (settembre 2026)"

IMPOSTAZIONI = Path.home() / ".orario_musicale.json"
# Helvetica non esiste su Windows: Segoe UI è il carattere di sistema
CARATTERE = "Segoe UI" if platform.system() == "Windows" else "Helvetica"
BINDINGS = (
    "single_select", "row_select", "column_select", "drag_select", "arrowkeys",
    "right_click_popup_menu", "rc_insert_row", "rc_delete_row", "copy", "cut", "paste",
    "delete", "undo", "edit_cell", "column_width_resize", "double_click_column_resize",
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
    return f"Risultati orario {datetime.now():%Y-%m-%d %H.%M}"


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
            self.lbl_stato.configure(text=f"File di una versione precedente: aggiunte {len(modifiche)} colonne/righe "
                                          "(Comune, Indirizzo, Civico, parametri trasporti). Premi Salva per tenerle.")

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

        self.tabs = ctk.CTkTabview(f, anchor="nw")
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
        if not hasattr(self, "btn_trasporti"):
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
        suggerimento = {
            FOGLIO_STUDENTI: "Doppio clic su una cella per modificarla. Le celle vuote in Docente 1/2 e KM vanno completate.",
            FOGLIO_DOCENTI: "Scrivi X nelle ore disponibili, A nelle ore di accompagnamento fissate, vuoto se non disponibile.",
        }.get(nome, "Doppio clic su una cella per modificarla. Tasto destro per inserire o eliminare righe.")
        ctk.CTkLabel(comandi, text=suggerimento, text_color="gray45").pack(side="left", padx=16)

        sheet = Sheet(parent, headers=list(t.intestazione), data=[list(r) for r in t.righe],
                      theme="light blue", show_row_index=True, header_font=(CARATTERE, 12, "bold"),
                      font=(CARATTERE, 12, "normal"))
        sheet.enable_bindings(*BINDINGS)
        sheet.pack(fill="both", expand=True)
        self._adatta_colonne(sheet, nome)
        self.fogli[nome] = sheet

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
        if self.schermata_dati is not None:
            self.schermata_dati.pack_forget()
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
            salva_tabelle(self.percorso, self.tabelle)
        except PermissionError:
            self.mostra_problemi([Problema("errore", self.percorso.name,
                                           "Il file è aperto in Excel: chiuderlo e salvare di nuovo.")], "Salvataggio non riuscito")
            return False
        except Exception as e:
            self.mostra_problemi([Problema("errore", self.percorso.name, f"Salvataggio non riuscito: {e}")],
                                 "Salvataggio non riuscito")
            return False
        self.lbl_stato.configure(text=f"Salvato alle {datetime.now():%H:%M}")
        return True

    # ── Calcolo ──
    def calcola(self) -> None:
        assert self.percorso is not None
        if not self.salva():
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
            intest = [f"{d.nome}\n{d.aula}" for d in docenti]
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
    def __init__(self, master, titolo: str, testo: str) -> None:
        super().__init__(master)
        self.title(titolo)
        self.geometry("900x600")
        self.transient(master)
        box = ctk.CTkTextbox(self, font=ctk.CTkFont(size=13), wrap="word")
        box.pack(fill="both", expand=True, padx=16, pady=16)
        box.insert("end", testo)
        box.configure(state="disabled")
        ctk.CTkButton(self, text="Chiudi", command=self.destroy).pack(pady=(0, 14))


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
