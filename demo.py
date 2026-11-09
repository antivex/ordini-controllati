# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "pypdf==6.14.2"]
# ///
"""Review local order documents and export reconciled item rows for ERP import."""

import argparse
import csv
import re
import tempfile
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from email import policy
from email.parser import BytesParser
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader
from pypdf.errors import PyPdfError
from profiles import Order, parse_profile


ROOT = Path(__file__).resolve().parent
LABELS = ("Ordine", "Data", "Azienda", "Prodotto", "Quantità", "Prezzo unitario EUR", "Totale EUR")


def field_value(label, raw):
    if label == "Data":
        for pattern, fmt in ((r"\d{4}-\d{2}-\d{2}", "%Y-%m-%d"), (r"\d{2}/\d{2}/\d{4}", "%d/%m/%Y")):
            if re.fullmatch(pattern, raw):
                parsed = datetime.strptime(raw, fmt).date()
                if parsed.year < 1900:
                    raise ValueError
                return parsed
        raise ValueError
    if label == "Quantità":
        if not re.fullmatch(r"[1-9][0-9]{0,5}", raw):
            raise ValueError
        return int(raw)
    if label.endswith(" EUR"):
        if not re.fullmatch(r"(?:0|[1-9][0-9]{0,7})(?:[.,][0-9]{2})?", raw):
            raise ValueError
        return Decimal(raw.replace(",", "."))
    if label == "Ordine" and not re.fullmatch(r"[A-Z0-9][A-Z0-9-]{0,39}", raw):
        raise ValueError
    if len(raw) > 120 or re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]", raw):
        raise ValueError
    return raw


def parse_order(filename, text):
    profiled = parse_profile(filename, text)
    if profiled is not None:
        return profiled
    order = Order(filename)
    if not text.strip():
        order.reasons.append("Testo assente: verificare il documento originale.")
        return order
    labels = {label.casefold(): label for label in LABELS}
    collected = defaultdict(list)
    for line in text.splitlines():
        key, sep, value = line.partition(":")
        label = labels.get(key.strip().casefold())
        if sep and label:
            collected[label].append(value.strip())
    if not collected:
        order.reasons.append("Profilo fornitore o formato non riconosciuto.")
        return order
    for label in LABELS:
        raw = collected[label]
        if not raw or not all(raw):
            order.reasons.append(f"{label}: campo mancante o vuoto.")
        elif len(set(raw)) > 1:
            order.reasons.append(f"{label}: valori discordanti nel documento.")
        else:
            try:
                order.values[label] = field_value(label, raw[0])
            except ValueError:
                order.reasons.append(f"{label}: formato o valore non valido.")
    quantity, price, total = (order.values.get(key) for key in LABELS[-3:])
    if None not in (quantity, price, total) and quantity * price != total:
        order.reasons.append("Totale dichiarato diverso da quantità × prezzo unitario.")
    return order


def read_order(path):
    try:
        if path.suffix.lower() == ".pdf":
            reader = PdfReader(path)
            if reader.is_encrypted:
                return Order(path.name, reasons=["PDF protetto: revisione manuale."])
            pages = [page.extract_text(extraction_mode="layout") or "" for page in reader.pages]
            order = parse_order(path.name, "\n".join(pages))
            if any(not page.strip() for page in pages) and any(page.strip() for page in pages):
                order.reasons.append("Almeno una pagina è priva di testo estraibile.")
            return order
        if path.suffix.lower() == ".eml":
            message = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
            bodies = [part for part in message.walk() if part.get_content_type() == "text/plain" and not part.is_attachment()]
            if len(bodies) != 1 or any(message.iter_attachments()):
                return Order(path.name, reasons=["Email senza corpo testuale univoco o con allegati da verificare."])
            return parse_order(path.name, bodies[0].get_content())
        return parse_order(path.name, path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, PyPdfError, ValueError, LookupError):
        return Order(path.name, reasons=["Documento illeggibile o danneggiato."])


def load_orders(folder):
    if not folder.is_dir():
        raise ValueError("La cartella di ingresso non esiste.")
    paths = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".txt", ".pdf", ".eml"})
    if not paths:
        raise ValueError("La cartella non contiene documenti TXT, EML o PDF.")
    orders = [read_order(path) for path in paths]
    by_id = defaultdict(list)
    for order in orders:
        company, number = (order.values.get(key) for key in ("Azienda", "Ordine"))
        if company and number:
            by_id[(company.casefold(), number)].append(order)
    for group in by_id.values():
        if len(group) > 1:
            for order in group:
                order.reasons.append("Numero ordine ripetuto per la stessa azienda.")
    return orders


def append_values(sheet, values):
    sheet.append(values)
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str):
            # Treat all source text as literal, including strings starting with '='.
            cell.data_type = "s"
        elif isinstance(cell.value, date):
            cell.number_format = "yyyy-mm-dd"


def format_sheet(sheet, widths):
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.row_dimensions[1].height = 28
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="15424C")
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center")


def write_workbook(orders, output):
    book = Workbook()
    book.properties.creator = book.properties.lastModifiedBy = "antivex"
    main = book.active
    main.title = "Ordini"
    details = book.create_sheet("Righe")
    review = book.create_sheet("Da controllare")
    main.append(["Documento", "Ordine", "Data", "Azienda", "Righe lette", "Somma righe EUR", "Totale dichiarato EUR", "Esito", "Righe nel CSV"])
    details.append(["Documento", "Ordine", "Codice", "Descrizione", "Quantità", "Prezzo EUR", "Riga dichiarata EUR", "Riga calcolata EUR", "Importazione"])
    review.append(["Documento", "Ordine", "Motivo"])
    for order in orders:
        status = "Da controllare" if order.reasons else "Controlli superati"
        v = order.values
        number = v.get("Ordine")
        calculated = sum(i.net for i in order.items) if order.items else None
        count = len(order.items)
        disposition = "Inclusa" if order.can_export else "Trattenuta"
        for item in order.items:
            append_values(details, [order.filename, number, item.sku, item.product, item.quantity, item.price, item.stated_net, item.net, disposition])
        if not order.items and v.get("Prodotto") is not None:
            count = 1
            if v.get("Quantità") is not None and v.get("Prezzo unitario EUR") is not None:
                calculated = v["Quantità"] * v["Prezzo unitario EUR"]
            append_values(details, [order.filename, number, None, v["Prodotto"], v.get("Quantità"), v.get("Prezzo unitario EUR"), None, calculated, "Archivio v1"])
        append_values(main, [order.filename, number, v.get("Data"), v.get("Azienda"), count, calculated, v.get("Totale EUR"), status, count if order.can_export else 0])
        for column in ("F", "G"):
            main[f"{column}{main.max_row}"].number_format = '#,##0.00 "€"'
        if order.reasons:
            main[f"H{main.max_row}"].fill = PatternFill("solid", fgColor="FBE6C1")
            reason = "\n".join(order.reasons)
            append_values(review, [order.filename, order.values.get("Ordine"), reason])
            review[f"C{review.max_row}"].alignment = Alignment(wrap_text=True, vertical="top")
            review.row_dimensions[review.max_row].height = 18 * len(order.reasons) + 8
    for row in details.iter_rows(min_row=2, min_col=6, max_col=8):
        for cell in row:
            cell.number_format = '#,##0.00 "€"'
    format_sheet(main, [29, 17, 16, 26, 15, 22, 25, 23, 17])
    format_sheet(details, [29, 17, 14, 25, 15, 18, 24, 23, 18])
    format_sheet(review, [29, 17, 94])
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temp:
        pending = Path(temp.name)
    try:
        book.save(pending)
        pending.replace(output)
    finally:
        pending.unlink(missing_ok=True)
        book.close()


CSV_HEADER = ["fornitore", "ordine", "data", "riga", "sku", "descrizione", "quantita", "prezzo_unitario", "importo_riga", "valuta"]


def write_outputs(orders, output):
    """Stage both files together; an earlier CSV must never describe a newer workbook."""
    output.parent.mkdir(parents=True, exist_ok=True)
    target_csv = output.with_suffix(".csv")
    lock = output.with_suffix(".export-lock")
    with lock.open("x") as owner:
        owner.write("Esportazione in corso")
    try:
        with tempfile.TemporaryDirectory(dir=output.parent) as directory:
            stage = Path(directory)
            staged_book, staged_csv = stage / "new.xlsx", stage / "new.csv"
            write_workbook(orders, staged_book)
            with staged_csv.open("w", encoding="utf-8", newline="") as stream:
                writer = csv.writer(stream, delimiter=";")
                writer.writerow(CSV_HEADER)
                for order in orders:
                    if order.can_export:
                        for pos, item in enumerate(order.items, 1):
                            writer.writerow([order.values["Azienda"], order.values["Ordine"], order.values["Data"].isoformat(), pos, item.sku, item.product, item.quantity, f"{item.price:.2f}", f"{item.net:.2f}", "EUR"])
            backup_csv = stage / "old.csv"
            if target_csv.exists():
                target_csv.replace(backup_csv)
            try:
                staged_book.replace(output)
            except OSError:
                if backup_csv.exists():
                    backup_csv.replace(target_csv)
                raise
            staged_csv.replace(target_csv)
    finally:
        lock.unlink()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=ROOT / "samples")
    parser.add_argument("--output", type=Path, default=ROOT / "output" / "ordini.xlsx")
    args = parser.parse_args()
    if args.output.suffix.lower() != ".xlsx":
        parser.error("L'uscita deve avere estensione .xlsx")
    try:
        orders = load_orders(args.input)
        write_outputs(orders, args.output)
    except (ValueError, OSError) as error:
        parser.exit(1, f"Errore: {error}\n")
    count = sum(len(o.items) for o in orders if o.can_export)
    print(f"{len(orders)} documenti | {sum(bool(o.reasons) for o in orders)} da controllare | {count} righe CSV | {args.output.name}, {args.output.with_suffix('.csv').name}")


if __name__ == "__main__":
    main()
