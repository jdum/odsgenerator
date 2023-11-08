from pathlib import Path

from odfdo import Document

from odsgenerator import odsgenerator as og

DATA = Path(__file__).parent / "data"
FILE_TUTO = DATA / "tutorial.json"


def make_ods_document(path):
    dest = path / "tutorial.ods"
    dest.unlink(missing_ok=True)
    og.file_to_ods(FILE_TUTO, dest)
    document = Document(dest)
    return document


def document_tables(path):
    document = make_ods_document(path)
    tables = document.body.get_tables()
    return tables


def test_tables(tmp_path):
    tables = document_tables(tmp_path)
    assert len(tables) == 7


def test_t0_name(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    assert table.name == "Tab 1"


def test_t0_rows(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    assert len(rows) == 2


def test_t0_r0_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[0]
    values = row.get_values()
    assert values[0].startswith("Minimal tab")
