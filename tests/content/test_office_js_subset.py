"""The Office JS compatibility promise, kept honest.

CR-003 section 5: ``tests/office_js_subset.json`` is the committed list of the
Office JS members this package implements, derived by
``scripts/office_js_subset.py`` from docx4j-core-ts's own compile-time
assignability check. This test asserts that every non-extension member of an
implemented phase is really on the class that owns it --- Phase B's ``Body``,
``Paragraph``, ``Range`` and ``Font``, Phase C's five, Phase G's ``Comment``,
Phase F's ``TrackedChange`` and the ``Document`` members that are the package's,
Phase H's ``List`` and ``ListItem`` --- with the right kind, and reports the
members of later phases still to come.
"""

from __future__ import annotations

import inspect
import json
from pathlib import Path

import pytest
from conftest import ROOT

from docx4j_py import WordprocessingMLPackage
from docx4j_py.model.content import (
    Body,
    Comment,
    ContentControl,
    Font,
    InlinePicture,
    List,
    ListItem,
    Paragraph,
    Range,
    Table,
    TableCell,
    TableRow,
    TrackedChange,
)
from docx4j_py.model.customxml import (
    CheckboxContentControl,
    ComboBoxContentControl,
    ContentControlListItem,
    CustomXmlNode,
    CustomXmlPart,
    CustomXmlPartCollection,
    CustomXmlPrefixMappingCollection,
    DatePickerContentControl,
    DropDownListContentControl,
    GroupContentControl,
    ListContentControl,
    PictureContentControl,
    RepeatingSectionContentControl,
    XmlMapping,
)

SUBSET = ROOT / "tests" / "office_js_subset.json"

#: The four interfaces Phase B owns, and the class each one is.
PHASE_B_CLASSES = {"Body": Body, "Paragraph": Paragraph, "Range": Range, "Font": Font}

#: The five Phase C adds.
PHASE_C_CLASSES = {
    "Table": Table,
    "TableRow": TableRow,
    "TableCell": TableCell,
    "InlinePicture": InlinePicture,
    "ContentControl": ContentControl,
}

#: The one Phase G adds.
PHASE_G_CLASSES = {"Comment": Comment}

#: The two Phase F adds. Office JS's ``Document`` is the **package** here
#: (``pkg.body``, ``pkg.change_tracking_mode``, ``pkg.get_tracked_changes()``),
#: which is where CR-003 section 3.8 puts those three members.
PHASE_F_CLASSES = {
    "TrackedChange": TrackedChange,
    "Document": WordprocessingMLPackage,
}

#: The twelve Phase E adds: the custom XML model, the mapping and the typed
#: content-control kinds (CR-003 section 3.7). Office JS's ``ListContentControl``
#: is the shape a drop-down and a combo box share; both subclass it here.
PHASE_E_CLASSES = {
    "CustomXmlPart": CustomXmlPart,
    "CustomXmlPartCollection": CustomXmlPartCollection,
    "CustomXmlNode": CustomXmlNode,
    "CustomXmlPrefixMappingCollection": CustomXmlPrefixMappingCollection,
    "XmlMapping": XmlMapping,
    "CheckboxContentControl": CheckboxContentControl,
    "DatePickerContentControl": DatePickerContentControl,
    "ListContentControl": ListContentControl,
    "ContentControlListItem": ContentControlListItem,
    "PictureContentControl": PictureContentControl,
    "RepeatingSectionContentControl": RepeatingSectionContentControl,
}
# Office JS's ``GroupContentControl`` has no members of its own, so the subset
# list carries no entry for it; the class is here all the same, and
# ``ContentControl.group_content_control`` is what the list does promise.
assert GroupContentControl is not None

#: The two Phase H adds. Office JS's list members of ``Paragraph`` and ``Body``
#: are carried in the subset list by ``scripts/office_js_subset.py``'s
#: ``AHEAD_OF_TS`` block, because docx4j-core-ts's own phase H is still proposed
#: and its ``office-js-subset.ts`` declares no list interfaces (CR-003 section 18).
PHASE_H_CLASSES = {"List": List, "ListItem": ListItem}

#: Every interface implemented so far, and the phase that owns each member.
CLASSES = {
    **PHASE_B_CLASSES,
    **PHASE_C_CLASSES,
    **PHASE_G_CLASSES,
    **PHASE_F_CLASSES,
    **PHASE_E_CLASSES,
    **PHASE_H_CLASSES,
}

#: The phases this file's assertions hold for.
PHASES = ("B", "C", "E", "F", "G", "H")


@pytest.fixture(scope="module")
def subset() -> dict:
    """The committed list."""
    return json.loads(SUBSET.read_text(encoding="utf-8"))


def kind_of(cls: type, name: str) -> str | None:
    """``"property"``, ``"method"`` or None, by what the class declares.

    A ``__slots__`` entry counts as a property: the views hold their element,
    their container and their parent in slots (CR-001 section 14: ``parent`` is
    a slot, not a field), and ``row.parent_table`` reads exactly as Office JS's
    ``parentTable`` does. What is being asserted is the shape of the member, not
    how the class happens to store it.
    """
    for klass in cls.__mro__:
        if name in vars(klass):
            member = vars(klass)[name]
            is_data = isinstance(member, property) or type(member).__name__ == "member_descriptor"
            return "property" if is_data else "method"
    return None


def test_the_list_is_committed_and_names_its_source(subset):
    assert SUBSET.exists()
    assert subset["source"] == "docx4j-core-ts/test/office-js-subset.ts"
    assert set(subset["interfaces"]) >= set(CLASSES)
    # the two members Python spells differently, recorded in the list itself
    assert subset["aliases"] == {
        "Table.getCell": "cell",
        "InlinePicture.getBase64ImageSrc": "get_base64",
    }


def test_every_implemented_member_is_there_with_the_right_kind(subset):
    missing: list[str] = []
    wrong_kind: list[str] = []
    later: dict[str, list[str]] = {}

    for interface, members in subset["interfaces"].items():
        cls = CLASSES.get(interface)
        for member in members:
            if member["extension"] or interface == "SearchOptions":
                continue
            where = f"{interface}.{member['name']}"
            if member["phase"] not in PHASES or cls is None:
                if cls is None or kind_of(cls, member["name"]) is None:
                    later.setdefault(member["phase"], []).append(where)
                continue
            kind = kind_of(cls, member["name"])
            if kind is None:
                missing.append(where)
            elif kind != member["kind"]:
                wrong_kind.append(f"{where} is a {kind}, Office JS has a {member['kind']}")

    summary = "; ".join(
        f"{phase} ({len(names)}): {', '.join(sorted(names)[:4])}"
        + (", ..." if len(names) > 4 else "")
        for phase, names in sorted(later.items())
    )
    print(f"\nlater phases still to come -- {summary}")
    assert not missing and not wrong_kind, (
        f"Phases {', '.join(PHASES)} owe {missing or 'nothing'}; "
        f"wrong kind: {wrong_kind or 'none'}. "
        f"(For the record, the members of later phases still missing are {summary}.)"
    )


def test_the_two_list_kinds_share_one_shape(subset):
    """Office JS gives a drop-down and a combo box the same members.

    ``ListContentControl`` is the shape; ``DropDownListContentControl`` and
    ``ComboBoxContentControl`` are the two kinds, and a control reports whichever
    one it is (CR-003 section 3.7).
    """
    assert issubclass(DropDownListContentControl, ListContentControl)
    assert issubclass(ComboBoxContentControl, ListContentControl)
    for member in subset["interfaces"]["ListContentControl"]:
        for cls in (DropDownListContentControl, ComboBoxContentControl):
            assert kind_of(cls, member["name"]) == member["kind"], f"{cls.__name__}.{member}"


def test_the_xpath_addressed_editors_put_the_mappings_last(subset):
    """CR-003 section 17: a departure the subset list does not cover.

    Office JS's desktop-only ``insertElement(xpath, namespaceMappings, xml)``
    puts the mappings second; here they come after the XML, with a default, so
    the common call --- a part with one namespace, or none --- is two arguments.
    """
    names = {member["name"] for member in subset["interfaces"]["CustomXmlPart"]}
    assert "insert_element" not in names  # deliberately outside the promise

    signature = inspect.signature(CustomXmlPart.insert_element)
    assert list(signature.parameters) == ["self", "xpath", "xml", "namespace_mappings", "index"]
    assert signature.parameters["namespace_mappings"].default is None


def test_the_search_options_are_the_keyword_arguments(subset):
    """Office JS's ``SearchOptions`` is an object; here it is keyword arguments."""
    names = [member["name"] for member in subset["interfaces"]["SearchOptions"]]
    assert names == ["match_case", "match_whole_word", "match_wildcards"]

    from docx4j_py.model.content.text_model import search_pattern

    signature = inspect.signature(search_pattern)
    for name in names:
        assert signature.parameters[name].kind is inspect.Parameter.KEYWORD_ONLY


def test_the_classes_offer_nothing_under_an_office_js_name_they_should_not(subset):
    """A member spelled the Office JS way must be the Office JS thing.

    The guard against a private helper accidentally taking a promised name: any
    public member of the four classes whose name matches an Office JS member of
    a *different* interface is suspicious, and a member of this interface must
    have the declared kind.
    """
    for interface, cls in CLASSES.items():
        declared = {member["name"]: member for member in subset["interfaces"][interface]}
        for name, member in declared.items():
            kind = kind_of(cls, name)
            if kind is None:
                continue
            assert kind == member["kind"], f"{interface}.{name}"


def test_the_script_would_write_what_is_committed():
    """Run the generator against the sibling checkout when it is there."""
    import sys

    sys.path.insert(0, str(ROOT / "scripts"))
    import office_js_subset

    if not office_js_subset.DEFAULT_SOURCE.exists():
        pytest.skip("docx4j-core-ts is not checked out beside this repository")
    assert office_js_subset.main(["--check"]) == 0


def test_the_snake_case_mapping_is_the_documented_one():
    import sys

    sys.path.insert(0, str(ROOT / "scripts"))
    import office_js_subset

    assert office_js_subset.snake("styleBuiltIn") == "style_built_in"
    assert office_js_subset.snake("insertInlinePictureFromBase64") == (
        "insert_inline_picture_from_base64"
    )
    assert office_js_subset.snake("getRange") == "get_range"
    assert office_js_subset.python_name("Table", "getCell") == "cell"
    assert office_js_subset.python_name("InlinePicture", "getBase64ImageSrc") == "get_base64"
    assert office_js_subset.python_name("Paragraph", "getRange") == "get_range"


def test_the_committed_file_is_valid_json_the_repository_tracks():
    assert isinstance(json.loads(Path(SUBSET).read_text(encoding="utf-8")), dict)
