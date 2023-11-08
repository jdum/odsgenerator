from odfdo import Document

from odsgenerator import odsgenerator as og


def test_sample1(tmp_path):
    path = tmp_path / "sample1.ods"
    raw = og.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
    with open(path, "wb") as file:
        file.write(raw)
    assert path.is_file()
    document = Document(path)
    tables = document.body.get_tables()
    table = tables[0]
    rows = table.get_rows()
    row1 = rows[0]
    values1 = row1.get_values()
    row2 = rows[1]
    values2 = row2.get_values()
    assert values1 == ["a", "b", "c"]
    assert values2 == [10, 20, 30]
