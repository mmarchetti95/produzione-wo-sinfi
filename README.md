# Produzione WO SINFI

Board di controllo produzione dei rilievi Walkout (QField / QFieldCloud) per la commessa SINFI.

**Board:** https://mmarchetti95.github.io/produzione-wo-sinfi/
(è anche nel [portale Controllo Produzione Rilievi](https://mmarchetti95.github.io/-Controllo-Produzione-Rilievi/))

Per chi la usa non serve installare nulla: si apre il link e basta.

## Come funziona

```
QField (campo) ──sync──▶ QFieldCloud ──▶ GitHub Action ──▶ sito/dati.json ──▶ GitHub Pages (board)
```

1. La GitHub Action [`aggiorna.yml`](.github/workflows/aggiorna.yml) esegue `qfc_report.py`, che legge da
   QFieldCloud i *delta* (modifiche sincronizzate) degli ultimi 60 giorni e li somma per giorno, operatore,
   progetto e layer.
2. Il risultato (`sito/dati.json`) viene pubblicato su GitHub Pages insieme alla board `sito/index.html`.
3. La board legge `dati.json`, ricalcola viste/filtri/totali nel browser ed esporta in Excel (SheetJS).

L'Action parte:
- **da sola** ogni 30 minuti, lunedì–sabato, circa 6–21 ora italiana;
- **a ogni push** su `main` (quindi ogni modifica al codice è online per tutti in 1-2 minuti);
- **a mano**: scheda *Actions* → *Aggiorna board* → *Run workflow* (o bottone "Aggiorna ora" nella board).

## Struttura

| File | Cosa fa |
|---|---|
| `qfc_report.py` | Motore di calcolo: login QFieldCloud, download delta, km netti/nuovi, create/modificate/cancellate. Produce JSON (board) o Excel. |
| `sito/index.html` | La board (HTML/CSS/JS in un solo file, nessuna build). |
| `.github/workflows/aggiorna.yml` | Aggiornamento automatico e pubblicazione su Pages. |
| `requirements.txt` | Librerie Python usate dall'Action. |

## Configurazione (una volta sola)

Credenziali QFieldCloud in **Settings → Secrets and variables → Actions → Secrets**:

| Secret | Valore |
|---|---|
| `QFC_URL` | `https://qfield.eagleprojects.cloud` |
| `QFC_USER` | utente di reporting |
| `QFC_PASSWORD` | password |

**Settings → Pages → Source: GitHub Actions.**

Le credenziali non devono mai finire nel codice: il file `.env` è escluso da `.gitignore`.

## Uso in locale (facoltativo)

```bash
pip install -r requirements.txt
cp .env.example .env        # e compila utente/password
python qfc_report.py --giorni 60 --json sito/dati.json   # dati per la board
python -m http.server -d sito 8000                        # poi apri http://localhost:8000
python qfc_report.py --dal 2026-10-01 --al 2026-10-12     # report Excel completo
```

## Note

- **I dati sono pubblici**: chiunque abbia il link vede operatori (utenti QFieldCloud), progetti e produzione.
- Se la password QFieldCloud cambia, aggiornare il secret `QFC_PASSWORD`: la board mostra un avviso se i dati
  sono più vecchi di 36 ore.
- GitHub sospende le Action programmate nei repository senza attività da 60 giorni: in quel caso basta
  riattivarle dalla scheda *Actions* (o fare un commit qualsiasi).
