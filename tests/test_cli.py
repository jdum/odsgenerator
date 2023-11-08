import re
import subprocess
from pathlib import Path

from odfdo import Document

from odsgenerator.cli import check_odfdo_version

RE_VERS = re.compile(r' *version *= *"(\S+)"$')
DATA = Path(__file__).parent / "data"
FILE2 = DATA / "test_minimal.json"


def read_proj_version():
    pyproject = Path("__file__").parent.parent / "pyproject.toml"
    with open(pyproject) as content:
        for line in content:
            if group := RE_VERS.match(line):
                return group[1]
    raise ValueError


def capture(command):
    proc = subprocess.Popen(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    out, err = proc.communicate()
    return out.strip(), err.strip(), proc.returncode


def test_odfdo_version():
    assert check_odfdo_version()


def test_no_param():
    command = ["odsgenerator"]
    out, err, exitcode = capture(command)
    assert exitcode == 2
    assert out == b""
    assert err.startswith(b"usage: odsgenerator [-h] [--version]")


def test_version():
    command = ["odsgenerator", "--version"]
    out, err, exitcode = capture(command)
    assert exitcode == 0
    expected = f"odsgenerator {read_proj_version()}".encode()
    assert out == expected
    assert err == b""


def test_help():
    command = ["odsgenerator", "--help"]
    out, err, exitcode = capture(command)
    assert exitcode == 0
    assert err == b""
    assert b"odsgenerator, an .ods generator" in out
    assert b"Usage" in out
    assert b"Arguments" in out
    assert b"Principle" in out
    assert b"Styles" in out
    assert b"decimal6_grid_06pt" in out


def test_generate(tmp_path):
    dest = tmp_path / "document.ods"
    dest.unlink(missing_ok=True)
    command = ["odsgenerator", str(FILE2), str(dest)]
    out, err, exitcode = capture(command)
    assert exitcode == 0
    assert err == b""
    assert out == b""
    assert dest.is_file()
    document = Document(dest)
    assert "spreadsheet" in document.container.mimetype
    tables = document.body.get_tables()
    assert len(tables) == 2
