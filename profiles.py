"""Readers for three fictional suppliers. Unconsumed document text requires review."""

import re
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal


@dataclass
class Item:
    sku: str
    product: str
    quantity: int
    price: Decimal
    stated_net: Decimal | None

    @property
    def net(self):
        return self.quantity * self.price


@dataclass
class Order:
    filename: str
    values: dict = field(default_factory=dict)
    reasons: list[str] = field(default_factory=list)
    items: list[Item] = field(default_factory=list)

    @property
    def can_export(self):
        return bool(self.items) and not self.reasons


PRICE = r"(?:0|[1-9][0-9]{0,7})(?:[.,][0-9]{2})?"
SKU = r"[A-Z][A-Z0-9]{1,8}-[0-9]{1,7}"
PRODUCT = r"[^\W\d_][\w '’/-]{0,95}"
QTY = r"[1-9][0-9]{0,5}"
CATALOG = {
    "tasselli bruma": ("CO-41", "Tasselli bruma"),
    "staffe onda": ("CO-52", "Staffe onda"),
}


def amount(text):
    return Decimal(text.replace(",", "."))


def set_date(order, text, fmt):
    try:
        day = datetime.strptime(text, fmt).date()
        if day.year < 1900:
            raise ValueError
        order.values["Data"] = day
    except ValueError:
        order.reasons.append("Data non valida.")


def add_item(order, sku, name, quantity, price, net=None):
    item = Item(sku, name, int(quantity), amount(price), amount(net) if net is not None else None)
    order.items.append(item)
    if item.stated_net is not None and item.net != item.stated_net:
        order.reasons.append(f"{sku}: importo di riga diverso da quantità × prezzo.")


def take_header(order, lines, pattern, message):
    if lines:
        match = re.fullmatch(pattern, lines[0])
        if match:
            lines.pop(0)
            return match.groups()
    order.reasons.append(message)
    return ()


def table_pdf(order, lines):
    head = take_header(order, lines, r"Ordine (RF-[0-9]+) del (\d{2}/\d{2}/\d{4})", "Intestazione ordine non riconosciuta.")
    if head:
        order.values["Ordine"] = head[0]
        set_date(order, head[1], "%d/%m/%Y")
    take_header(order, lines, r"Codice\s+Descrizione\s+Pezzi\s+Prezzo EUR\s+Riga EUR", "Colonne della tabella non riconosciute.")
    for pos, line in enumerate(lines, 1):
        cells = re.split(r"\s{2,}", line)
        if len(cells) == 5 and all(re.fullmatch(p, val) for p, val in zip((SKU, PRODUCT, QTY, PRICE, PRICE), cells)):
            add_item(order, *cells)
        else:
            order.reasons.append(f"Riga tabella {pos}: codice, descrizione, quantità o importi da verificare.")


def block_pdf(order, lines):
    head = take_header(order, lines, r"Conferma di acquisto (VO-[0-9]+)", "Identificativo conferma non riconosciuto.")
    if head:
        order.values["Ordine"] = head[0]
    day = take_header(order, lines, r"Emessa il (\d{4}-\d{2}-\d{2})", "Data della conferma assente.")
    if day:
        set_date(order, day[0], "%Y-%m-%d")
    index = 0
    while index < len(lines):
        head = re.fullmatch(rf"Art\. ({SKU}) / ({PRODUCT})", lines[index])
        index += 1
        if not head:
            order.reasons.append(f"Testo inatteso nella sezione articoli, riga {index}.")
            continue
        detail = re.fullmatch(rf"({QTY}) pezzi x EUR ({PRICE}) = EUR ({PRICE})", lines[index]) if index < len(lines) else None
        if not detail:
            order.reasons.append(f"{head[1]}: dettaglio quantità/prezzo/importo assente o illeggibile.")
            continue
        index += 1
        add_item(order, *head.groups(), *detail.groups())


def read_profile_pdf(order, text, company):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if lines and lines[-1] == "Documento sintetico":
        lines.pop()
    take_header(order, lines, re.escape(company), "Intestazione del fornitore inattesa.")
    is_table = company == "Ramofuso Demo"
    total_pattern = rf"Totale ordine EUR ({PRICE})" if is_table else rf"Importo complessivo:\s*({PRICE}) EUR"
    total = re.fullmatch(total_pattern, lines[-1]) if lines else None
    if total:
        order.values["Totale EUR"] = amount(total[1])
        lines.pop()
    else:
        order.reasons.append("Totale documento assente o ambiguo.")
    (table_pdf if is_table else block_pdf)(order, lines)


def prose_email(order, text):
    body = " ".join(text.split()).replace("’", "'")
    pattern = (r"(?:Buongiorno|Salve),\s*per l'ordine (CO-[0-9]+) del (\d{2}/\d{2}/\d{4}) "
               r"(?:vi ordiniamo|vorremmo ordinare) (.+?)\.\s*"
               rf"Il totale concordato è di ({PRICE}) euro\.\s*(?:Grazie|Cordiali saluti),\s*Cartaonda Demo")
    match = re.fullmatch(pattern, body, re.IGNORECASE)
    if not match:
        order.reasons.append("Email: riferimento, articoli, totale o testo aggiuntivo da verificare.")
        return
    number, day, clause, total = match.groups()
    order.values["Ordine"] = number.upper()
    set_date(order, day, "%d/%m/%Y")
    order.values["Totale EUR"] = amount(total)
    for pos, phrase in enumerate(re.split(r"\s+e\s+|,\s+(?=\d+\s)", clause, flags=re.IGNORECASE), 1):
        item = re.fullmatch(rf"({QTY}) (.+?) a ({PRICE}) euro(?: ciascun[oa])?", phrase, re.IGNORECASE)
        if not item:
            order.reasons.append(f"Articolo {pos}: quantità o prezzo non esplicito nella mail.")
            continue
        product = CATALOG.get(item[2].casefold())
        if product is None:
            order.reasons.append(f"Articolo {pos}: descrizione fuori dal catalogo del fornitore.")
            continue
        add_item(order, *product, item[1], item[3])


def parse_profile(filename, text):
    companies = [name for name in ("Ramofuso Demo", "Officina Velombra Demo", "Cartaonda Demo") if name.casefold() in text.casefold()]
    if not companies:
        return None
    order = Order(filename)
    if len(companies) != 1:
        order.reasons.append("Il documento richiama più profili fornitore.")
        return order
    company = companies[0]
    order.values["Azienda"] = company
    if company == "Cartaonda Demo":
        prose_email(order, text)
    else:
        read_profile_pdf(order, text, company)
    if not order.items:
        order.reasons.append("Nessuna riga articolo verificabile.")
    if not order.reasons:
        total = sum(item.net for item in order.items)
        if total != order.values["Totale EUR"]:
            order.reasons.append(f"Somma righe {total:.2f} EUR diversa dal totale documento {order.values['Totale EUR']:.2f} EUR.")
    return order
