import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from demo import ROOT, load_orders, parse_order, read_order, write_workbook


SOURCE = (ROOT / "samples" / "01-ordine.txt").read_text(encoding="utf-8")


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
                self.assertEqual(book["Ordini"]["E2"].value, "=3+4")
                self.assertEqual(book["Ordini"]["E2"].data_type, "s")
            finally:
                book.close()

    def test_end_to_end_including_real_pdf_extraction(self):
        orders = load_orders(ROOT / "samples")
        self.assertEqual(len(orders), 6)
        self.assertEqual([o.filename[:2] for o in orders if o.reasons], ["05", "06"])
        self.assertEqual(orders[2].values["Ordine"], "OR-103")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "orders.xlsx"
            write_workbook(orders, path)
            write_workbook(orders, path)
            book = load_workbook(path)
            try:
                self.assertEqual(book.sheetnames, ["Ordini", "Da controllare"])
                self.assertEqual(book["Ordini"].max_row, 7)
                self.assertEqual(book["Da controllare"].max_row, 3)
                self.assertIsNone(book["Ordini"]["F6"].value)
                self.assertIsNone(book["Ordini"]["H7"].value)
                self.assertIn("Quantità", book["Da controllare"]["C2"].value)
                self.assertIn("discordanti", book["Da controllare"]["C3"].value)
            finally:
                book.close()


if __name__ == "__main__":
    unittest.main()
