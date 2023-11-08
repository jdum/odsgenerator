import pytest

from odsgenerator import odsgenerator as og


def test_no_args_odsgen():
    with pytest.raises(TypeError):
        og.file_to_ods()


def test_one_args_odsgen():
    with pytest.raises(TypeError):
        og.file_to_ods("some file")
