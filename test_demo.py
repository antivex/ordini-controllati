import csv
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from demo import ROOT, load_orders, parse_order, read_order, write_outputs, write_workbook


SAMPLES = ROOT / "samples"
SOURCE = (SAMPLES / "legacy" / "01-ordine.txt").read_text(encoding="utf-8")


class DemoTests(unittest.TestCase):
    def test_valid_order_preserves_types(self):
        order = parse_order("one.txt", SOURCE)
        self.assertEqual(order.reasons, [])
        self.assertEqual(order.values["Data"], date(2026, 10, 12))
        self.assertEqual(order.values["Totale EUR"], Decimal("34.00"))

    def test_no_guess_for_conflict_or_missing_field(self):
        order = parse_order("one.txt", SOURCE.replace("Quantità: 8", "") + "\nTotale EUR: 50,00")
        self.assertNotIn("Quantità", order.values)
        self.assertNotIn("Totale EUR", order.values)
        self.assertEqual(len(order.reasons), 2)

    def test_arithmetic_error_retains_reported_amount(self):
        order = parse_order("one.txt", SOURCE.replace("34,00", "43,00"))
        self.assertTrue(any("diverso" in reason for reason in order.reasons))
        self.assertEqual(order.values["Totale EUR"], Decimal("43.00"))

    def test_invalid_dates_and_ambiguous_thousands(self):
        order = parse_order("one.txt", SOURCE.replace("2026-10-12", "31/02/2026").replace("34,00", "1,200"))
        self.assertNotIn("Data", order.values)
        self.assertNotIn("Totale EUR", order.values)

    def test_unwritable_control_character_goes_to_review(self):
        order = parse_order("one.txt", SOURCE.replace("Tassello nebbia", "Articolo\x00rotto"))
        self.assertNotIn("Prodotto", order.values)
        self.assertTrue(any("Prodotto" in reason for reason in order.reasons))

    def test_duplicate_id_flags_every_copy(self):
        with tempfile.TemporaryDirectory() as tmp:
            for filename in ("a.txt", "b.txt"):
                (Path(tmp) / filename).write_text(SOURCE, encoding="utf-8")
            orders = load_orders(Path(tmp))
            self.assertTrue(all(any("ripetuto" in reason for reason in order.reasons) for order in orders))

    def test_broken_pdf_counts_as_document(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "broken.pdf"
            path.write_bytes(b"%PDF-1.4\n")
            self.assertTrue(read_order(path).reasons)

    def test_untrusted_text_stays_literal(self):
        order = parse_order("one.txt", SOURCE.replace("Tassello nebbia", "=3+4"))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "orders.xlsx"
            write_workbook([order], path)
            book = load_workbook(path)
            try:
                self.assertEqual(book["Righe"]["D2"].value, "=3+4")
                self.assertEqual(book["Righe"]["D2"].data_type, "s")
            finally:
                book.close()

    def test_end_to_end_including_real_pdf_extraction(self):
        orders = load_orders(SAMPLES / "legacy")
        self.assertEqual(len(orders), 6)
        self.assertEqual([o.filename[:2] for o in orders if o.reasons], ["05", "06"])
        self.assertEqual(orders[2].values["Ordine"], "OR-103")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "orders.xlsx"
            write_workbook(orders, path)
            write_workbook(orders, path)
            book = load_workbook(path)
            try:
                self.assertEqual(book.sheetnames, ["Ordini", "Righe", "Da controllare"])
                self.assertEqual(book["Ordini"].max_row, 7)
                self.assertEqual(book["Da controllare"].max_row, 3)
                self.assertIsNone(book["Righe"]["E6"].value)
                self.assertIsNone(book["Ordini"]["G7"].value)
                self.assertIn("Quantità", book["Da controllare"]["C2"].value)
                self.assertIn("discordanti", book["Da controllare"]["C3"].value)
            finally:
                book.close()

    def test_three_profiles_reconcile_and_export_only_complete_orders(self):
        orders = load_orders(SAMPLES)
        by_number = {o.values["Ordine"]: o for o in orders}
        self.assertEqual([len(o.items) for o in orders], [2, 3, 2, 2, 2, 1])
        self.assertEqual(sum(i.net for i in by_number["CO-301"].items), Decimal("49.00"))
        self.assertIn("29.50", by_number["VO-202"].reasons[0])
        self.assertIn("27.50", by_number["VO-202"].reasons[0])
        self.assertIn("Articolo 2", by_number["CO-302"].reasons[0])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "ordini.xlsx"
            write_outputs(orders, path)
            with path.with_suffix(".csv").open(encoding="utf-8", newline="") as f:
                rows = list(csv.DictReader(f, delimiter=";"))
            self.assertEqual(len(rows), 9)
            self.assertEqual({r["ordine"] for r in rows}, {"RF-101","RF-102","VO-201","CO-301"})
            self.assertEqual(sum(Decimal(r["importo_riga"]) for r in rows), Decimal("177.50"))
            self.assertEqual([(r["riga"],r["sku"]) for r in rows if r["ordine"]=="CO-301"], [("1","CO-41"),("2","CO-52")])
            book = load_workbook(path, read_only=True)
            try:
                self.assertEqual([s.max_row for s in book], [7,13,3])
                self.assertEqual([book["Da controllare"][cell].value for cell in ("B2","B3")], ["VO-202","CO-302"])
            finally:
                book.close()
            write_outputs([by_number["VO-202"]], path)
            self.assertEqual(len(path.with_suffix(".csv").read_text().splitlines()), 1)

    def test_malformed_table_line_blocks_whole_order_even_when_total_matches(self):
        from pypdf import PdfReader
        text = PdfReader(SAMPLES / "01-ramofuso.pdf").pages[0].extract_text(extraction_mode="layout")
        order = parse_order("table.pdf", text.replace("Totale ordine EUR", "extra sconosciuto\nTotale ordine EUR"))
        self.assertFalse(order.can_export)
        self.assertTrue(any("Riga tabella" in r for r in order.reasons))

    def test_stated_line_net_is_checked_independently(self):
        from pypdf import PdfReader
        text = PdfReader(SAMPLES / "03-velombra.pdf").pages[0].extract_text(extraction_mode="layout")
        order = parse_order("conferma.pdf", text.replace("= EUR 19,00", "= EUR 18,00"))
        self.assertFalse(order.can_export)
        self.assertIn("VE-11", order.reasons[0])

    def test_email_outside_catalog_and_extra_instructions_are_held(self):
        body = (SAMPLES / "05-cartaonda.eml").read_text(encoding="utf-8").split("\n\n", 1)[1]
        for change in (body.replace("tasselli bruma", "pezzi ignoti"), body + " Applicare sconto 10%.", body + " Ramofuso Demo"):
            self.assertFalse(parse_order("mail.eml", change).can_export)

    def test_duplicate_profiled_order_is_never_exported(self):
        with tempfile.TemporaryDirectory() as tmp:
            for name in ("first.eml","second.eml"):
                (Path(tmp) / name).write_bytes((SAMPLES / "05-cartaonda.eml").read_bytes())
            self.assertTrue(all(not o.can_export for o in load_orders(Path(tmp))))

    def test_readers_ignore_filenames_when_selecting_profiles(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "other-name.pdf"
            path.write_bytes((SAMPLES / "02-ramofuso.pdf").read_bytes())
            self.assertEqual(read_order(path).values["Ordine"], "RF-102")


if __name__ == "__main__":
    unittest.main()
