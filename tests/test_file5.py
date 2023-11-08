from pathlib import Path

from odfdo import Document

from odsgenerator import odsgenerator as og

DATA = Path(__file__).parent / "data"
FILE5 = DATA / "test_formula.json"


def make_ods_document(path):
    dest = path / "test_formula.ods"
    dest.unlink(missing_ok=True)
    og.file_to_ods(FILE5, dest)
    document = Document(dest)
    return document


def document_tables(path):
    document = make_ods_document(path)
    tables = document.body.get_tables()
    return tables


def test_tables(tmp_path):
    tables = document_tables(tmp_path)
    assert len(tables) == 1


def test_t0_r3_formula(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[3]
    cell = row.get_cell(1)
    expected = "of:=[.B3]+[.B2]"
    assert cell.formula == expected
