"""
Crea il file Excel di input `input_orario.xlsx` (più fogli da compilare),
precompilandolo con i dati che si possono leggere dai file esistenti:

  - classi e gruppi 2026-27 ... .docx   → studenti, strumenti, docenti
  - Recapiti musicale 2025.26 ... .xls   → km casa-scuola (incrocio per cognome+nome)

Uso:
    python crea_template.py

Il file prodotto va poi completato a mano (disponibilità docenti, gruppi LMC,
docenti della classe 1ª, km dei nuovi iscritti).
"""

from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

import docx
import xlrd
from orario.template import costruisci_workbook

ROOT = Path(__file__).parent
DOCX = next(ROOT.glob("classi e gruppi*.docx"))
XLS = next(ROOT.glob("Recapiti musicale*.xls"))
OUT = ROOT / "input_orario.xlsx"

GIORNI = ["Lun", "Mar", "Mer", "Gio", "Ven"]
ORE = ["13:30", "14:30", "15:30", "16:30"]
FASCE = [f"{g} {o}" for g in GIORNI for o in ORE]  # 20 colonne

# Aule dell'anno scorso (dal PDF 2025-26), solo come valore di partenza.
AULE_2025_26 = {
    "Vitangeli": "1ALM", "Lala": "2° p A", "Bini": "5ALM", "Pacini M": "M0",
    "Donnini": "M1", "Bartolozzi": "M2", "Francini": "M3", "Salaris": "M4",
    "Paganin": "M5", "Corsini": "M6", "Pacini C": "4ALM", "Paccosi": "M7",
    "Simonelli": "M8", "Lissoni": "3ALM", "Innocenti": "3ALM", "Mazzei": "M9",
    "Perc 2": "M 10", "Di Bonaventura": "M 11", "Caselli": "2ALM",
    "Buonaccorsi": "2ALM",
}


# ── utilità ──────────────────────────────────────────────────────────────────

def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("’", "'").upper()
    return re.sub(r"\s+", " ", s).strip()


def celle_uniche(row) -> list[str]:
    """Le celle unite in Word vengono ripetute: teniamo una sola copia per <tc>."""
    out, visti = [], set()
    for c in row.cells:
        if id(c._tc) in visti:
            continue
        visti.add(id(c._tc))
        out.append(c.text.strip())
    return out


def pulisci_docente(d: str) -> str:
    d = re.sub(r"\s+", " ", d or "").strip()
    return {"FL Lala": "Lala", "Pacini": "Pacini M"}.get(d, d)


# ── lettura docx ─────────────────────────────────────────────────────────────

def leggi_studenti(doc) -> list[dict]:
    studenti = []
    for classe, tabella in enumerate(doc.tables[:5], start=1):
        for row in tabella.rows:
            c = celle_uniche(row)
            if len(c) < 4 or not c[1] or c[1].upper() in {"COGNOME", ""}:
                continue
            if c[0].upper().startswith("CLASSE"):
                continue
            cognome, nome = c[1], c[2]
            if not cognome or not nome:
                continue
            strum1 = c[3] if len(c) > 3 else ""
            doc1 = c[4] if len(c) > 4 else ""
            strum2 = c[5] if len(c) > 5 else ""
            doc2 = c[6] if len(c) > 6 else ""
            if classe == 5:
                strum2, doc2 = "", ""
            studenti.append({
                "classe": classe,
                "cognome": cognome.replace("*", "").strip(),
                "nome": nome.replace("*", "").strip(),
                "strum1": strum1.upper(), "doc1": pulisci_docente(doc1),
                "strum2": strum2.upper(), "doc2": pulisci_docente(doc2),
            })
    return studenti


# ── lettura xls (km) ─────────────────────────────────────────────────────────

def spezza_civico(indirizzo: str) -> tuple[str, str]:
    """'Via Corsini, 515' → ('Via Corsini', '515'); 'VIA CASA PASQUINETTI 22/C' → ('Via Casa Pasquinetti', '22/C')."""
    t = re.sub(r"\s+", " ", (indirizzo or "").strip().strip(",.")).strip()
    m = re.match(r"^(.*?)[,\s]+(?:n\.?\s*|nr\.?\s*|n°\s*)?(\d+\s*[/\-]?\s*[A-Za-z0-9]{0,3})\s*$", t, re.IGNORECASE)
    if m and m.group(1):
        return m.group(1).strip(" ,.").title(), m.group(2).replace(" ", "").rstrip("/-").upper()
    return t.strip(" ,.").title(), ""


def leggi_km(path: Path) -> dict[tuple[str, str], dict]:
    """(COGNOME, NOME) → {km, comune, indirizzo, civico} dai recapiti."""
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    header = [norm(str(c.value)) for c in sh.row(0)]
    i_cog, i_nome, i_km = header.index("COGNOME"), header.index("NOME"), header.index("KM")
    i_com = next((i for i, h in enumerate(header) if h.startswith("COMUNE")), None)
    i_ind = next((i for i, h in enumerate(header) if h.startswith("INDIRIZZO")), None)
    km = {}
    for r in range(1, sh.nrows):
        cog = norm(str(sh.cell_value(r, i_cog)))
        nome = norm(str(sh.cell_value(r, i_nome)))
        if not cog:
            continue
        try:
            val = float(sh.cell_value(r, i_km))
        except (TypeError, ValueError):
            val = None
        via, civ = spezza_civico(str(sh.cell_value(r, i_ind))) if i_ind is not None else ("", "")
        km[(cog, nome)] = {"km": val, "comune": str(sh.cell_value(r, i_com)).strip().title() if i_com is not None else "",
                           "indirizzo": via, "civico": civ}
    return km


def cerca_km(km: dict, cognome: str, nome: str) -> dict | None:
    cog, nom = norm(cognome), norm(nome)
    if (cog, nom) in km:
        return km[(cog, nom)]
    # nome abbreviato nel docx ("ELENA" vs "ELENA ALEKSEEVNA", "MARIO H.")
    primo = nom.split(" ")[0].rstrip(".")
    candidati = [v for (c, n), v in km.items() if c == cog and n.split(" ")[0] == primo]
    if len(candidati) == 1:
        return candidati[0]
    candidati = [v for (c, n), v in km.items() if c == cog]
    if len(candidati) == 1:
        return candidati[0]
    return None


def main() -> None:
    doc = docx.Document(DOCX)
    studenti = leggi_studenti(doc)
    km = leggi_km(XLS)

    senza_km = []
    for s in studenti:
        rec = cerca_km(km, s["cognome"], s["nome"]) or {}
        s["km"] = rec.get("km")
        s["comune"], s["indirizzo"], s["civico"] = rec.get("comune", ""), rec.get("indirizzo", ""), rec.get("civico", "")
        if s["km"] is None:
            senza_km.append(s)

    docenti = sorted({d for s in studenti for d in (s["doc1"], s["doc2"]) if d})
    strumenti_per_docente: dict[str, set[str]] = {}
    for s in studenti:
        for strum, d in ((s["strum1"], s["doc1"]), (s["strum2"], s["doc2"])):
            if d:
                strumenti_per_docente.setdefault(d, set()).add(strum)

    righe_docenti = [{"nome": d, "strumenti": ", ".join(sorted(strumenti_per_docente[d])),
                      "aula": AULE_2025_26.get(d, ""), "ore_accomp": 0,
                      "note": "Cattedra Percussioni 2: nome del docente da definire" if d == "Perc 2" else ""}
                     for d in docenti]
    wb = costruisci_workbook(studenti, righe_docenti, esempi=True)
    wb.save(OUT)

    # ── riepilogo a video ──
    print(f"Creato: {OUT.name}")
    print(f"Studenti letti dal Word: {len(studenti)}  (per classe: "
          + ", ".join(f"{c}ª={sum(1 for s in studenti if s['classe']==c)}" for c in range(1, 6)) + ")")
    print(f"Docenti trovati: {len(docenti)} → {', '.join(docenti)}")
    print(f"Studenti senza KM (da completare, evidenziati in giallo): {len(senza_km)}")
    for s in senza_km:
        print(f"   - {s['classe']}ª {s['cognome']} {s['nome']}")
    senza_doc = [s for s in studenti if not s["doc1"] or (s["classe"] != 5 and not s["doc2"])]
    print(f"Studenti senza docente (da completare, evidenziati in giallo): {len(senza_doc)}")


if __name__ == "__main__":
    sys.exit(main())
