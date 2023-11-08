from odsgenerator import odsgenerator as og


def test_style_list():
    assert len(og.DEFAULT_STYLES) == 37


def test_style_list2():
    for style in og.DEFAULT_STYLES:
        assert style["name"]
        assert style["definition"]


def test_style_list3():
    known = set()
    for style in og.DEFAULT_STYLES:
        assert style["name"] not in known
        known.add(style["name"])


def test_style_list4():
    known = set()
    for style in og.DEFAULT_STYLES:
        assert style["definition"] not in known
        known.add(style["definition"])
