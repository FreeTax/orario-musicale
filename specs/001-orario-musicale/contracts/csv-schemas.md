# CSV Schemas — Input

Tutti i file CSV usano `;` come separatore e UTF-8 senza BOM come codifica.
La riga 1 è sempre l'intestazione. Celle vuote indicano campo assente (non errore).

---

## studenti.csv

| Colonna | Tipo | Obbligatorio | Valori validi | Note |
|---------|------|:------------:|---------------|------|
| cognome | text | ✅ | Non vuoto | Usato come identificatore interno |
| nome | text | ✅ | Non vuoto | |
| classe | integer | ✅ | 1–5 | |
| km | decimal | ✗ | ≥ 0 | Default 0.0 se assente o non numerico |
| docente_1str | text | ✅ | Non vuoto | Deve corrispondere a un docente in disponibilita_docenti.csv |
| strumento_1 | text | ✅ | Non vuoto | |
| docente_2str | text | ✗ | Non vuoto se presente | Vuoto per classe 5; deve corrispondere a un docente se presente |
| strumento_2 | text | ✗ | Non vuoto se presente | Vuoto per classe 5 |
| ore_consecutive | text | ✅ | SI / NO | Case-insensitive |
| giorno_unico | text | ✅ | SI / NO | Case-insensitive |

**Esempio**:
```
cognome;nome;classe;km;docente_1str;strumento_1;docente_2str;strumento_2;ore_consecutive;giorno_unico
Rossi;Mario;3;12.5;Bianchi;Violino;Verdi;Teoria;NO;NO
Neri;Sara;5;3;Ferrari;Pianoforte;;;NO;NO
Conti;Luca;1;8;Bianchi;Violino;Verdi;Solfeggio;SI;NO
```

---

## gruppi_lmc.csv

| Colonna | Tipo | Obbligatorio | Valori validi | Note |
|---------|------|:------------:|---------------|------|
| numero_gruppo | integer | ✅ | > 0, univoco | |
| docente | text | ✅ | Non vuoto | Deve corrispondere a un docente in disponibilita_docenti.csv |
| studente_1 | text | ✅ | Cognome esistente | Solo studenti di classi 3–5 |
| studente_2 | text | ✅ | Cognome esistente | Solo studenti di classi 3–5 |
| studente_3 | text | ✗ | Cognome esistente se presente | Vuoto per gruppi da 2 studenti |
| studente_4 | text | ✗ | Cognome esistente se presente | Vuoto per gruppi da 2-3 studenti |

**Esempio**:
```
numero_gruppo;docente;studente_1;studente_2;studente_3;studente_4
1;Verdi;Rossi;Bruni;Galli;
2;Ferrari;Neri;Conti;;
```

---

## disponibilita_docenti.csv

| Colonna | Tipo | Obbligatorio | Valori validi | Note |
|---------|------|:------------:|---------------|------|
| docente | text | ✅ | Non vuoto, univoco | |
| lun_1330 | text | ✗ | "X" o vuoto | X = disponibile; vuoto = non disponibile |
| lun_1430 | text | ✗ | "X" o vuoto | |
| lun_1530 | text | ✗ | "X" o vuoto | |
| lun_1630 | text | ✗ | "X" o vuoto | |
| mar_1330 | text | ✗ | "X" o vuoto | |
| mar_1430 | text | ✗ | "X" o vuoto | |
| mar_1530 | text | ✗ | "X" o vuoto | |
| mar_1630 | text | ✗ | "X" o vuoto | |
| mer_1330 | text | ✗ | "X" o vuoto | |
| mer_1430 | text | ✗ | "X" o vuoto | |
| mer_1530 | text | ✗ | "X" o vuoto | |
| mer_1630 | text | ✗ | "X" o vuoto | |
| gio_1330 | text | ✗ | "X" o vuoto | |
| gio_1430 | text | ✗ | "X" o vuoto | |
| gio_1530 | text | ✗ | "X" o vuoto | |
| gio_1630 | text | ✗ | "X" o vuoto | |
| ven_1330 | text | ✗ | "X" o vuoto | |
| ven_1430 | text | ✗ | "X" o vuoto | |
| ven_1530 | text | ✗ | "X" o vuoto | |
| ven_1630 | text | ✗ | "X" o vuoto | |

**Ordine colonne obbligatorio**: le 20 colonne devono apparire nell'ordine
lun→ven, 1330→1630 come sopra.

**Esempio**:
```
docente;lun_1330;lun_1430;lun_1530;lun_1630;mar_1330;mar_1430;mar_1530;mar_1630;mer_1330;mer_1430;mer_1530;mer_1630;gio_1330;gio_1430;gio_1530;gio_1630;ven_1330;ven_1430;ven_1530;ven_1630
Bianchi;X;X;X;X;X;X;X;X;;;;X;X;X;;X;X;X;X
Ferrari;;X;X;X;;X;X;X;X;X;X;X;;X;X;X;;X;X;X
Verdi;X;X;;X;X;X;;X;X;X;;X;X;;X;X;X;X;;X
```
