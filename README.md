# Dagli ordini a Excel · Orders to Excel

Sei ordini inventati entrano come testo di email salvato in UTF-8 o come PDF con testo selezionabile. Lo script produce un Excel con una riga per documento. Il foglio **Da controllare** raccoglie i documenti dubbi e il motivo.

Six fictional orders arrive as saved email text or text-based PDFs. The script writes one row per document to Excel. The **Da controllare** (Needs review) sheet lists the exceptions and explains each one.

![Prima e dopo / Before and after: un PDF senza quantità e la corrispondente segnalazione in Excel](docs/prima-dopo.png)

[Apri l’Excel di esempio / Open the sample workbook](docs/ordini-esempio.xlsx)

## Esecuzione · Run

Con [uv](https://docs.astral.sh/uv/) installato, dalla cartella del progetto eseguire il comando qui sotto. From the project folder, run the command below with uv installed. The first run may download Python dependencies. Document processing runs locally.

```sh
uv run demo.py
```

Con Python 3.10+ già disponibile / With an existing Python 3.10+ environment:

```sh
python -m pip install -r requirements.txt
python demo.py
```

Aprire **output/ordini.xlsx**. I sei esempi in **samples/** danno quattro righe con controlli superati e due da rivedere. Nel PDF OR-105 manca la quantità. Nel testo OR-106 compaiono due totali diversi. Le celle incerte restano vuote. Se un totale dichiarato è univoco ma non torna con quantità × prezzo, il foglio conserva il valore originale e segnala la differenza.

Open **output/ordini.xlsx**. Four samples pass the checks. OR-105 has no quantity and OR-106 contains conflicting totals. Uncertain fields stay blank. When a single stated total disagrees with quantity × unit price, the workbook preserves the stated amount and flags the discrepancy.

## Formato e limiti · Input and scope

Un articolo per ordine. Usare le sette etichette italiane presenti nei campioni, ciascuna su una riga. Sono accettati importi in euro senza separatori delle migliaia e date ISO o GG/MM/AAAA. Lo script verifica anche gli identificativi ripetuti per la stessa azienda. I campi sconosciuti vengono ignorati. Nessun accesso alla posta, OCR, interpretazione di prosa libera, riconciliazione IVA/sconti o invio a un gestionale.

Use the seven Italian field labels in the samples, one field per line and one product per order. Amounts are in EUR without thousands separators. Dates use ISO or DD/MM/YYYY. Duplicate order IDs for the same company are flagged. Unknown labels are ignored. Image-only PDFs need human review. Inbox access, OCR, free-form language interpretation, tax/discount reconciliation and ERP posting are outside this demo.

Percorsi alternativi e test / Custom paths and tests:

```sh
uv run demo.py --input samples --output output/orders.xlsx
uv run --with-requirements requirements.txt python -m unittest -v
```

Con le dipendenze già installate basta `python -m unittest -v`. With dependencies installed, run `python -m unittest -v`.

Autore / Author: **antivex**. Realizzato con il supporto dell’AI, su dati interamente sintetici. Built with AI assistance using entirely synthetic data. [MIT License](LICENSE).
