from decimal import Decimal

from odfdo import Document

from odsgenerator import odsgenerator as og


def test_sample1(tmp_path):
    path = tmp_path / "sample2.ods"
    raw = og.ods_bytes(
        [
            {
                "name": "table tab",
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
    assert values2 == [Decimal("10.00"), Decimal("20.00"), Decimal("30.00")]
