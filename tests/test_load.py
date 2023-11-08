from pathlib import Path

from odfdo import Document

from odsgenerator import odsgenerator as og

DATA = Path(__file__).parent / "data"
FILE1 = DATA / "test_json.json"
FILE2 = DATA / "test_minimal.json"
FILE3 = DATA / "test_yaml.yml"
FILE4 = DATA / "test_use_case.json"
FILE5 = DATA / "test_formula.json"

FILES = (FILE1, FILE2, FILE3, FILE4, FILE5)


def test_run(tmp_path):
    for file in FILES:
        output = tmp_path / (file.stem + "1.ods")
        og.file_to_ods(file, output)
        assert output.is_file()


def test_load(tmp_path):
    for file in FILES:
        output = tmp_path / (file.stem + "2.ods")
        og.file_to_ods(file, output)
        doc = Document(output)
        assert isinstance(doc, Document)
