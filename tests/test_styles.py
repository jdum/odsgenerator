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


def test_new_style_keys_default_and_common():
    auto_def = (
        '<style:style style:family="table-cell" style:parent-style-name="Default">'
        '<style:text-properties fo:font-weight="bold"/></style:style>'
    )
    default_def = (
        '<style:default-style style:family="table-cell">'
        '<style:table-cell-properties style:decimal-places="2"/></style:default-style>'
    )
    common_def = (
        '<style:style style:family="table-cell">'
        '<style:text-properties fo:color="#ff0000"/></style:style>'
    )
    content = {
        "body": [{"name": "Tab 1", "table": [["a", "b"], [10, 20]]}],
        "styles": [{"name": "auto1", "definition": auto_def}],
        "default_styles": [{"definition": default_def}],
        "common_styles": [{"name": "CustomCommon", "definition": common_def}],
    }
    gen = og.ODSGenerator(content)
    # Check default style is in office:styles
    office_styles = gen.doc.styles.get_element("//office:styles")
    defaults = office_styles.get_elements("style:default-style")
    assert len(defaults) >= 1
    # Check common style CustomCommon is in office:styles
    assert gen.doc.get_style("table-cell", "CustomCommon") is not None


def test_backward_compatibility_without_new_keys():
    auto_def = (
        '<style:style style:family="table-cell" style:parent-style-name="Default">'
        '<style:text-properties fo:font-weight="bold"/></style:style>'
    )
    content = {
        "body": [{"name": "Tab 1", "table": [["a", "b"], [10, 20]]}],
        "styles": [{"name": "auto1", "definition": auto_def}],
    }
    gen = og.ODSGenerator(content)
    assert gen.doc.body is not None
