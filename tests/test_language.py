import io
import json

import pytest
from odfdo import Document

from odsgenerator import odsgenerator as og


def test_language_default():
    raw = og.ods_bytes([[["a", "b"]]])
    doc = Document(io.BytesIO(raw))
    assert doc.language == "fr-FR"
    assert doc.meta.language == "fr-FR"


def test_language_dict():
    content = {
        "language": "en-US",
        "body": [[["Hello", "World"]]],
    }
    raw = og.ods_bytes(content)
    doc = Document(io.BytesIO(raw))
    assert doc.language == "en-US"
    assert doc.meta.language == "en-US"
    assert doc.styles.default_language == "en-US"


def test_language_dict_override():
    content = {
        "language": "fr-FR",
        "body": [[["Hello", "World"]]],
    }
    raw = og.ods_bytes(content, language="en-US")
    doc = Document(io.BytesIO(raw))
    assert doc.language == "en-US"
    assert doc.meta.language == "en-US"


def test_language_list_with_param():
    content = [[["Alpha", "Beta"]]]
    raw = og.ods_bytes(content, language="de-DE")
    doc = Document(io.BytesIO(raw))
    assert doc.language == "de-DE"
    assert doc.meta.language == "de-DE"


def test_language_content_to_ods(tmp_path):
    dest = tmp_path / "test_lang.ods"
    content = {
        "language": "es-ES",
        "body": [[["Hola", "Mundo"]]],
    }
    og.content_to_ods(content, dest)
    doc = Document(dest)
    assert doc.language == "es-ES"
    assert doc.meta.language == "es-ES"


def test_language_file_to_ods_json(tmp_path):
    json_path = tmp_path / "data.json"
    dest = tmp_path / "output.ods"
    data = {
        "language": "en-GB",
        "body": [[["Colour", "Flavor"]]],
    }
    json_path.write_text(json.dumps(data), encoding="utf8")
    og.file_to_ods(json_path, dest)
    doc = Document(dest)
    assert doc.language == "en-GB"
    assert doc.meta.language == "en-GB"


def test_language_file_to_ods_override(tmp_path):
    json_path = tmp_path / "data.json"
    dest = tmp_path / "output.ods"
    data = {
        "language": "fr-FR",
        "body": [[["Data"]]],
    }
    json_path.write_text(json.dumps(data), encoding="utf8")
    og.file_to_ods(json_path, dest, language="pt-BR")
    doc = Document(dest)
    assert doc.language == "pt-BR"
    assert doc.meta.language == "pt-BR"


def test_language_file_to_ods_yaml(tmp_path):
    yaml_path = tmp_path / "data.yml"
    dest = tmp_path / "output.ods"
    yaml_content = "language: it-IT\nbody:\n  - - [Ciao, Mondo]\n"
    yaml_path.write_text(yaml_content, encoding="utf8")
    og.file_to_ods(yaml_path, dest)
    doc = Document(dest)
    assert doc.language == "it-IT"
    assert doc.meta.language == "it-IT"


def test_language_invalid():
    content = {
        "language": "invalid_code_format",
        "body": [[["a"]]],
    }
    with pytest.raises(TypeError):
        og.ods_bytes(content)
