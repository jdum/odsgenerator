#!/usr/bin/env python

import unittest
from unittest import TestCase

import os
from odfdo import Document, Table, Row, Cell
import odsgenerator.odsgenerator as og
from decimal import Decimal

FILE1 = "test_json.json"
FILE2 = "test_minimal.json"
FILE3 = "test_yaml.yml"
FILE4 = "test_use_case.json"


class TestMain(TestCase):
    def test_no_args_odsgen(self):
        with self.assertRaises(TypeError) as cm:
            og.file_to_ods()
            self.assertEqual(cm.exception.code, 0)

    def test_one_args_odsgen(self):
        with self.assertRaises(TypeError) as cm:
            og.file_to_ods("some file")
            self.assertEqual(cm.exception.code, 0)


class TestLoadFiles(TestCase):
    def setUp(self):
        self.files = FILE1, FILE2, FILE3
        for f in self.files:
            output = f + ".ods"
            if os.path.isfile(output):
                os.remove(output)

    def tearDown(self):
        for f in self.files:
            output = f + ".ods"
            try:
                os.remove(output)
            except IOError:
                pass

    def test_run(self):
        for f in self.files:
            output = f + ".ods"
            og.file_to_ods(f, output)
            self.assertTrue(os.path.isfile(output))

    def test_load(self):
        for f in self.files:
            output = f + ".ods"
            og.file_to_ods(f, output)
            d = Document(output)
            self.assertIsInstance(d, Document)


class TestFile1(TestCase):
    def setUp(self):
        self.output = FILE1 + ".ods"
        if os.path.isfile(self.output):
            os.remove(self.output)
        og.file_to_ods(FILE1, self.output)
        self.document = Document(self.output)
        self.body = self.document.body

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_tables(self):
        tables = self.body.get_tables()
        self.assertEqual(len(tables), 2)

    def test_t0_name(self):
        tables = self.body.get_tables()
        t = tables[0]
        self.assertEqual(t.name, "first tab")

    def test_t0_rows(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t0_r0_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(
            values, ["spanned cell", None, None, "d", "e", "f", "g", "h", "i", "j"]
        )

    def test_t0_r1_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(values, [None, None, None, 30, 40, 50, 60, 70, 80, 90])

    def test_t0_r2_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [1, 11, 21, 31, 41, 51, 61, 71, 81, 91])

    def test_t0_r3_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(values, [2, 12, 22, 32, 42, 52, 62, 72, 82, 92])

    def test_t1_name(self):
        tables = self.body.get_tables()
        t = tables[1]
        self.assertEqual(t.name, "second tab")

    def test_t1_rows(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t1_r0_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(values, ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"])

    def test_t1_r1_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(
            values,
            [
                Decimal("100.01"),
                Decimal("110.02"),
                "hop",
                130,
                140,
                150,
                160,
                170,
                180,
                190,
            ],
        )

    def test_t1_r2_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [101, 111, 121, 131, 141, 151, 161, 171, 181, 191])

    def test_t1_r3_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(
            values, [Decimal("102.314"), 112, 122, 132, 0, 152, 0, 172, 0, 192]
        )


class TestFile2(TestCase):
    def setUp(self):
        self.output = FILE2 + ".ods"
        if os.path.isfile(self.output):
            os.remove(self.output)
        og.file_to_ods(FILE2, self.output)
        self.document = Document(self.output)
        self.body = self.document.body

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_tables(self):
        tables = self.body.get_tables()
        self.assertEqual(len(tables), 2)

    def test_t0_name(self):
        tables = self.body.get_tables()
        t = tables[0]
        self.assertEqual(t.name, "Tab 1")

    def test_t0_rows(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t0_r0_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(values, ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"])

    def test_t0_r1_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(values, [0, 10, 20, 30, 40, 50, 60, 70, 80, 90])

    def test_t0_r2_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [1, 11, 21, 31, 41, 51, 61, 71, 81, 91])

    def test_t0_r3_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(values, [2, 12, 22, 32, 42, 52, 62, 72, 82, 92])

    def test_t1_name(self):
        tables = self.body.get_tables()
        t = tables[1]
        self.assertEqual(t.name, "Tab 2")

    def test_t1_rows(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t1_r0_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(values, ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"])

    def test_t1_r1_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(
            values,
            [100, 110, 120, 130, 140, 150, 160, 170, 180, 190],
        )

    def test_t1_r2_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [101, 111, 121, 131, 141, 151, 161, 171, 181, 191])

    def test_t1_r3_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(values, [102, 112, 122, 132, 142, 152, 162, 172, 182, 192])


class TestFile3(TestCase):
    def setUp(self):
        self.output = FILE3 + ".ods"
        if os.path.isfile(self.output):
            os.remove(self.output)
        og.file_to_ods(FILE3, self.output)
        self.document = Document(self.output)
        self.body = self.document.body

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_tables(self):
        tables = self.body.get_tables()
        self.assertEqual(len(tables), 2)

    def test_t0_name(self):
        tables = self.body.get_tables()
        t = tables[0]
        self.assertEqual(t.name, "first tab")

    def test_t0_rows(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t0_r0_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(
            values, ["spanned cell", None, None, "d", "e", "f", "g", "h", "i", "j"]
        )

    def test_t0_r1_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(values, [None, None, None, 30, 40, 50, 60, 70, 80, 90])

    def test_t0_r2_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [1, 11, 21, 31, 41, 51, 61, 71, 81, 91])

    def test_t0_r3_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(values, [2, 12, 22, 32, 42, 52, 62, 72, 82, 92])

    def test_t1_name(self):
        tables = self.body.get_tables()
        t = tables[1]
        self.assertEqual(t.name, "second tab")

    def test_t1_rows(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        self.assertEqual(len(rows), 4)

    def test_t1_r0_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(values, ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"])

    def test_t1_r1_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[1]
        values = r.get_values()
        self.assertEqual(
            values,
            [
                Decimal("100.01"),
                Decimal("110.02"),
                "hop",
                130,
                140,
                150,
                160,
                170,
                180,
                190,
            ],
        )

    def test_t1_r2_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[2]
        values = r.get_values()
        self.assertEqual(values, [101, 111, 121, 131, 141, 151, 161, 171, 181, 191])

    def test_t1_r3_values(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        r = rows[3]
        values = r.get_values()
        self.assertEqual(values, ["abc", 112, 122, 132, 0, 152, 0, 172, 0, 192])


class TestFile4(TestCase):
    def setUp(self):
        self.output = FILE4 + ".ods"
        if os.path.isfile(self.output):
            os.remove(self.output)
        og.file_to_ods(FILE4, self.output)
        self.document = Document(self.output)
        self.body = self.document.body

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_tables(self):
        tables = self.body.get_tables()
        self.assertEqual(len(tables), 2)

    def test_t0_name(self):
        tables = self.body.get_tables()
        t = tables[0]
        self.assertEqual(t.name, "Results")

    def test_t0_rows(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        self.assertEqual(len(rows), 41)

    def test_t0_r0_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[0]
        values = r.get_values()
        self.assertEqual(
            values,
            [
                "Surface id",
                "Surface name",
                "Surface layer",
                "Surface area [m2]",
                "Average value [h/day]",
                "0h00",
                "0h30",
                "1h00",
                "1h30",
                "2h00",
                "2h30",
                "3h00",
                "3h30",
                "4h00",
                "4h30",
                "5h00",
                "5h30",
                "6h00",
                "6h30",
                "7h00",
                "7h30",
                "Grid",
                "Comments",
            ],
        )

    def test_t0_r15_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[15]
        values = r.get_values()
        self.assertEqual(
            values,
            [
                92393,
                "92393",
                "wall façade",
                Decimal("16.880013"),
                Decimal("6.77"),
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                100,
                0,
                0,
                "very detailed (approx 4 sensors per m²)",
                "",
            ],
        )

    def test_t0_r40_values(self):
        tables = self.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r = rows[40]
        values = r.get_values()
        self.assertEqual(
            values,
            [
                106612,
                "106612",
                "wall façade",
                Decimal("41.800048"),
                Decimal("0.89"),
                0,
                Decimal("69.14"),
                Decimal("30.86"),
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                "very detailed (approx 4 sensors per m²)",
                "",
            ],
        )

    def test_t1_name(self):
        tables = self.body.get_tables()
        t = tables[1]
        self.assertEqual(t.name, "Scale")

    def test_t1_rows(self):
        tables = self.body.get_tables()
        t = tables[1]
        rows = t.get_rows()
        self.assertEqual(len(rows), 17)


class Testsample1(TestCase):
    def setUp(self):
        self.output = "sample1.ods"
        if os.path.isfile(self.output):
            os.remove(self.output)

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_sample1(self):
        raw = og.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
        with open(self.output, "wb") as f:
            f.write(raw)
        self.assertTrue(os.path.isfile(self.output))
        doc = Document(self.output)
        tables = doc.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r1 = rows[0]
        values1 = r1.get_values()
        r2 = rows[1]
        values2 = r2.get_values()
        self.assertEqual(values1, ["a", "b", "c"])
        self.assertEqual(values2, [10, 20, 30])


class TestSample2(TestCase):
    def setUp(self):
        self.output = "sample2.ods"
        if os.path.isfile(self.output):
            os.remove(self.output)

    def tearDown(self):
        try:
            os.remove(self.output)
        except IOError:
            pass

    def test_sample2(self):
        raw = og.ods_bytes(
            [
                {
                    "name": "first tab",
                    "style": "cell_decimal2",
                    "table": [
                        {
                            "row": ["a", "b", "c"],
                            "style": "bold_center_bg_gray_grid_06pt",
                        },
                        [10, 20, 30],
                    ],
                }
            ]
        )
        with open(self.output, "wb") as f:
            f.write(raw)
        self.assertTrue(os.path.isfile(self.output))
        doc = Document(self.output)
        tables = doc.body.get_tables()
        t = tables[0]
        rows = t.get_rows()
        r1 = rows[0]
        values1 = r1.get_values()
        r2 = rows[1]
        values2 = r2.get_values()
        self.assertEqual(values1, ["a", "b", "c"])
        self.assertEqual(
            values2, [Decimal("10.00"), Decimal("20.00"), Decimal("30.00")]
        )


if __name__ == "__main__":
    unittest.main()
