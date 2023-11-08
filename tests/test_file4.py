from decimal import Decimal
from pathlib import Path

from odfdo import Document

from odsgenerator import odsgenerator as og

DATA = Path(__file__).parent / "data"
FILE4 = DATA / "test_use_case.json"


def make_ods_document(path):
    dest = path / "test_use_case.ods"
    dest.unlink(missing_ok=True)
    og.file_to_ods(FILE4, dest)
    document = Document(dest)
    return document


def document_tables(path):
    document = make_ods_document(path)
    tables = document.body.get_tables()
    return tables


def test_tables(tmp_path):
    tables = document_tables(tmp_path)
    assert len(tables) == 2


def test_t0_name(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    assert table.name == "Results"


def test_t0_rows(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    assert len(rows) == 41


def test_t0_r0_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[0]
    values = row.get_values()
    assert values == [
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
    ]


def test_t0_r15_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[15]
    values = row.get_values()
    assert values == [
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
    ]


def test_t0_r40_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[40]
    values = row.get_values()
    assert values == [
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
    ]


def test_t1_name(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    assert table.name == "Scale"


def test_t1_rows(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    assert len(rows) == 17
