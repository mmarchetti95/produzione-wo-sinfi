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
   QFieldCloud i *delta* (modifiche sincronizzate) degli ultimi 12 mesi e li somma per giorno, operatore,
   progetto e layer (con il tipo di geometria del layer: punto, linea, poligono). Il selettore *Periodo* della board lavora su tutto questo storico.
2. Il risultato (`sito/dati.json`) viene pubblicato su GitHub Pages insieme alla board `sito/index.html`.
3. La board legge `dati.json`, ricalcola viste/filtri/totali e il compenso stimato (km netti × 204 €/km, costante
   `TARIFFA_KM` in `sito/index.html`) nel browser ed esporta il CSV del report produzione.

L'Action parte:
- **da sola** ogni 30 minuti, lunedì–sabato, circa 6–21 ora italiana;
- **a ogni push** su `main` (quindi ogni modifica al codice è online per tutti in 1-2 minuti);
- **a mano**: pulsante **Aggiorna ora** nella board (avvia l'Action e ricarica da solo i dati quando ha finito),
  oppure scheda *Actions* → *Aggiorna board* → *Run workflow*.

## Struttura

| File | Cosa fa |
|---|---|
| `qfc_report.py` | Motore di calcolo: login QFieldCloud, download delta, km netti/nuovi (nei netti si sottraggono solo cancellazioni e accorciamenti di linee create da QField, la rete di base rimossa va in `km_base`), create/modificate/cancellate, punti creati e tipo di geometria di ogni layer, foto e orari di inizio/fine (prima e ultima foto, letti dal nome del file), orario dell'ultima sincronizzazione. Produce JSON (board) o Excel. |
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
| `BOARD_DISPATCH_TOKEN` | token GitHub del pulsante "Aggiorna ora" (vedi sotto) |

**Settings → Pages → Source: GitHub Actions.**

### Token del pulsante "Aggiorna ora"

1. GitHub → *Settings* (del proprio profilo) → *Developer settings* → *Personal access tokens* → **Fine-grained tokens** → *Generate new token*.
2. *Repository access*: **Only select repositories** → `produzione-wo-sinfi`.
3. *Permissions* → *Repository permissions* → **Actions: Read and write** (nient'altro).
4. Scadenza: la più lunga consentita; annotarsi la data e rinnovarlo prima.
5. Salvarlo nel secret `BOARD_DISPATCH_TOKEN` e rilanciare l'Action una volta.

L'Action scrive il token in `sito/config.js`, che esiste solo nel sito pubblicato (non nel repository: GitHub revoca
i token trovati nel codice dei repository pubblici). **Il token è leggibile da chiunque apra la board**: con quei
permessi permette solo di avviare, annullare o cancellare le esecuzioni dell'Action di questo repository, non di
leggere i secret né di modificare il codice. Se viene abusato basta revocarlo e crearne uno nuovo.
Senza il secret il pulsante torna a aprire la pagina dell'Action su GitHub.

Le credenziali non devono mai finire nel codice: il file `.env` è escluso da `.gitignore`.

## Uso in locale (facoltativo)

```bash
pip install -r requirements.txt
cp .env.example .env        # e compila utente/password
python qfc_report.py --giorni 365 --json sito/dati.json  # dati per la board
python -m http.server -d sito 8000                        # poi apri http://localhost:8000
python qfc_report.py --dal 2026-10-01 --al 2026-10-12     # report Excel completo
```

## Note

- **I dati sono pubblici**: chiunque abbia il link vede operatori (utenti QFieldCloud), progetti e produzione.
- Se la password QFieldCloud cambia, aggiornare il secret `QFC_PASSWORD`: la board mostra un avviso se i dati
  sono più vecchi di 36 ore.
- GitHub sospende le Action programmate nei repository senza attività da 60 giorni: in quel caso basta
  riattivarle dalla scheda *Actions* (o fare un commit qualsiasi).
