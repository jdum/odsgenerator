from pathlib import Path

from odfdo import Document

from odsgenerator import odsgenerator as og

DATA = Path(__file__).parent / "data"
FILE2 = DATA / "test_minimal.json"


def make_ods_document(path):
    dest = path / "test_minimal.ods"
    dest.unlink(missing_ok=True)
    og.file_to_ods(FILE2, dest)
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
    assert table.name == "Tab 1"


def test_t0_rows(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    assert len(rows) == 4


def test_t0_r0_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[0]
    values = row.get_values()
    assert values == ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]


def test_t0_r1_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[1]
    values = row.get_values()
    assert values == [0, 10, 20, 30, 40, 50, 60, 70, 80, 90]


def test_t0_r2_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[2]
    values = row.get_values()
    assert values == [1, 11, 21, 31, 41, 51, 61, 71, 81, 91]


def test_t0_r3_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[0]
    rows = table.get_rows()
    row = rows[3]
    values = row.get_values()
    assert values == [2, 12, 22, 32, 42, 52, 62, 72, 82, 92]


def test_t1_name(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    assert table.name == "Tab 2"


def test_t1_rows(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    assert len(rows) == 4


def test_t1_r0_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    row = rows[0]
    values = row.get_values()
    assert values == ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]


def test_t1_r1_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    row = rows[1]
    values = row.get_values()
    assert values == [100, 110, 120, 130, 140, 150, 160, 170, 180, 190]


def test_t1_r2_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    row = rows[2]
    values = row.get_values()
    assert values == [101, 111, 121, 131, 141, 151, 161, 171, 181, 191]


def test_t1_r3_values(tmp_path):
    tables = document_tables(tmp_path)
    table = tables[1]
    rows = table.get_rows()
    row = rows[3]
    values = row.get_values()
    assert values == [102, 112, 122, 132, 142, 152, 162, 172, 182, 192]
