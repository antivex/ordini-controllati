# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl==3.1.5", "pypdf==6.14.2"]
# ///
"""Turn labelled, single-item order documents into a local review workbook."""

import argparse
import re
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader
from pypdf.errors import PyPdfError


ROOT = Path(__file__).resolve().parent
LABELS = ("Ordine", "Data", "Azienda", "Prodotto", "Quantità", "Prezzo unitario EUR", "Totale EUR")


@dataclass
class Order:
    filename: str
    values: dict = field(default_factory=dict)
    reasons: list[str] = field(default_factory=list)


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
            pages = [page.extract_text() or "" for page in reader.pages]
            order = parse_order(path.name, "\n".join(pages))
            if any(not page.strip() for page in pages) and any(page.strip() for page in pages):
                order.reasons.append("Almeno una pagina è priva di testo estraibile.")
            return order
        return parse_order(path.name, path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, PyPdfError, ValueError, KeyError):
        return Order(path.name, reasons=["Documento illeggibile o danneggiato."])


def load_orders(folder):
    if not folder.is_dir():
        raise ValueError("La cartella di ingresso non esiste.")
    paths = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".txt", ".pdf"})
    if not paths:
        raise ValueError("La cartella non contiene documenti TXT o PDF.")
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
    review = book.create_sheet("Da controllare")
    main.append(["Documento", *LABELS, "Esito"])
    review.append(["Documento", "Ordine", "Motivo"])
    for order in orders:
        status = "Da controllare" if order.reasons else "Controlli superati"
        append_values(main, [order.filename, *(order.values.get(key) for key in LABELS), status])
        for column in ("G", "H"):
            main[f"{column}{main.max_row}"].number_format = '#,##0.00 "€"'
        if order.reasons:
            main[f"I{main.max_row}"].fill = PatternFill("solid", fgColor="FBE6C1")
            reason = "\n".join(order.reasons)
            append_values(review, [order.filename, order.values.get("Ordine"), reason])
            review[f"C{review.max_row}"].alignment = Alignment(wrap_text=True, vertical="top")
            review.row_dimensions[review.max_row].height = 18 * len(order.reasons) + 8
    format_sheet(main, [28, 17, 15, 25, 25, 13, 22, 18, 22])
    format_sheet(review, [28, 17, 70])
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temp:
        pending = Path(temp.name)
    try:
        book.save(pending)
        pending.replace(output)
    finally:
        pending.unlink(missing_ok=True)
        book.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=ROOT / "samples")
    parser.add_argument("--output", type=Path, default=ROOT / "output" / "ordini.xlsx")
    args = parser.parse_args()
    if args.output.suffix.lower() != ".xlsx":
        parser.error("L'uscita deve avere estensione .xlsx")
    try:
        orders = load_orders(args.input)
        write_workbook(orders, args.output)
    except (ValueError, OSError) as error:
        parser.exit(1, f"Errore: {error}\n")
    print(f"{len(orders)} documenti | {sum(bool(o.reasons) for o in orders)} da controllare | {args.output.name}")


if __name__ == "__main__":
    main()
