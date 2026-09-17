"""``pkg.compatibility_mode``: which Word a document says it is written for.

CR-003 section 3.4 (added 2026-09-17) and section 17.11. Three things are
pinned: the read is free and honest (12 when the document says nothing, which is
what Word assumes, and ``/word/settings.xml`` byte for byte), the write is a
recorded mutation of that part, and a verb that writes a Word 2013 feature into
an older document **warns** rather than changing the mode behind the caller.
"""

from __future__ import annotations

import pytest
from conftest import part_bytes, reloaded, sample

from docx4j_py.model.content import ContentError
from docx4j_py.model.content.compatibility import (
    COMPATIBILITY_MODES,
    DEFAULT_COMPATIBILITY_MODE,
    NEW_DOCUMENT_MODE,
    CompatibilityMode,
)

#: Word 2010 wrote this one: ``compatibilityMode`` 14.
WORD_2010 = "2010-sample1.docx"
#: Word 2013 or later wrote this one: 15.
WORD_2013 = "DrawingML_GraphicData_wps.docx"
#: This one has a ``w:compat`` with nothing in it at all.
NO_SETTING = "Images.docx"


# ---------------------------------------------------------------------------
# reading
# ---------------------------------------------------------------------------


def test_the_mode_is_read_from_the_settings_part():
    assert sample(WORD_2010).compatibility_mode == 14
    assert sample(WORD_2013).compatibility_mode == 15
    # no setting at all is Word 2007, which is what Word assumes and why it
    # shows "Compatibility Mode" in the title bar for such a document
    assert sample(NO_SETTING).compatibility_mode == DEFAULT_COMPATIBILITY_MODE == 12


def test_the_enum_beside_it_reads_the_same():
    package = sample(WORD_2010)

    assert package.compatibility_mode == CompatibilityMode.WORD_2010
    assert CompatibilityMode.WORD_2013 == 15  # and Word 2016, 2019 and 365
    assert COMPATIBILITY_MODES == (11, 12, 14, 15)
    assert NEW_DOCUMENT_MODE == 15


def test_reading_the_mode_leaves_the_settings_part_byte_for_byte():
    package = sample(WORD_2010)
    before = part_bytes(package.save())

    assert package.compatibility_mode == 14
    assert package.compatibility_mode == 14  # the cached read
    assert package.document_settings_part.is_unmarshalled is False

    after = part_bytes(package.save())
    assert after["word/settings.xml"] == before["word/settings.xml"]


def test_describe_reports_the_mode():
    description = sample(WORD_2010).describe()

    assert description.compatibility_mode == 14
    assert description.to_dict()["compatibility_mode"] == 14


# ---------------------------------------------------------------------------
# writing
# ---------------------------------------------------------------------------


def test_the_setter_writes_the_setting_and_records_it():
    package = sample(WORD_2010)
    package.changes.clear()

    package.compatibility_mode = 15

    assert package.compatibility_mode == 15
    change = package.last_change
    assert change.operation == "compatibility_mode"
    assert "/word/settings.xml" in change.parts_touched
    assert change.text_after == "15"

    # the value is replaced, not added beside the one that was there
    saved = reloaded(package)
    settings = saved.document_settings_part.contents
    modes = [c.val for c in settings.compat.compat_setting if c.name == "compatibilityMode"]
    assert modes == ["15"]
    assert saved.compatibility_mode == 15


def test_the_setter_adds_the_setting_to_a_document_that_has_none():
    package = sample(NO_SETTING)

    package.compatibility_mode = CompatibilityMode.WORD_2013

    saved = reloaded(package)
    assert saved.compatibility_mode == 15
    settings = saved.document_settings_part.contents
    assert [c.name for c in settings.compat.compat_setting] == ["compatibilityMode"]


def test_a_mode_word_does_not_use_is_refused():
    package = sample(WORD_2010)

    with pytest.raises(ContentError) as error:
        package.compatibility_mode = 16  # there is no 16: Word 2016 writes 15

    assert error.value.code == "compatibility.mode_invalid"
    assert "15" in error.value.hint
    assert package.compatibility_mode == 14


def test_a_dry_run_of_the_setter_leaves_the_real_package_alone():
    package = sample(WORD_2010)
    package.body.paragraphs  # noqa: B018 - read the parts first, as a caller would
    package.document_settings_part.contents  # noqa: B018 - CR-003 section 12.5
    before = package.save()

    with package.dry_run() as trial:
        trial.compatibility_mode = 15
        assert trial.compatibility_mode == 15

    assert package.save() == before
    assert package.compatibility_mode == 14


def test_a_dry_run_over_a_settings_part_nobody_read_still_leaves_the_mode():
    """CR-003 section 12.5's limit: the trial's first look unmarshals the part.

    The bytes of that one part are then the re-marshalled ones --- which is true
    of every trial, not of this setter --- and what has to hold is that the mode
    itself, and every other part, are untouched.
    """
    package = sample(WORD_2010)
    package.body.paragraphs  # noqa: B018
    before = part_bytes(package.save())

    with package.dry_run() as trial:
        trial.compatibility_mode = 15

    after = part_bytes(package.save())
    assert package.compatibility_mode == 14
    assert b'w:name="compatibilityMode" w:uri="' in after["word/settings.xml"]
    assert b'w:val="15"' not in after["word/settings.xml"]
    assert {name: data for name, data in after.items() if name != "word/settings.xml"} == {
        name: data for name, data in before.items() if name != "word/settings.xml"
    }


# ---------------------------------------------------------------------------
# the warnings (CR-003 section 17.11)
# ---------------------------------------------------------------------------


def warnings_of(package) -> tuple[str, ...]:
    """The warnings of the last call, joined for a substring check."""
    return package.last_change.warnings


def test_a_word_2013_feature_warns_on_an_older_document():
    package = sample(WORD_2010)
    hit = package.body.find("first")[0]
    comment = hit.range(package.body).insert_comment("why")

    package.changes.clear()
    comment.resolved = True
    assert any("w15:done" in warning for warning in warnings_of(package))
    assert any("pkg.compatibility_mode = 15" in warning for warning in warnings_of(package))

    package.changes.clear()
    control = package.body.paragraphs[0].insert_content_control("RepeatingSection")
    assert any("w15:repeatingSection" in warning for warning in warnings_of(package))

    package.changes.clear()
    control.color = "#FF0000"
    assert any("w15:color" in warning for warning in warnings_of(package))

    package.changes.clear()
    control.appearance = "Tags"
    assert any("w15:appearance" in warning for warning in warnings_of(package))

    # and the mode is not changed for the caller: that is the caller's decision
    assert package.compatibility_mode == 14


def test_a_repeating_section_item_and_a_w15_binding_warn_too():
    package = sample("invoice2013.docx")  # Word 2013 markup, mode 14
    section = next(c for c in package.body.content_controls if c.type == "RepeatingSection")

    package.changes.clear()
    section.repeating_section_content_control.insert_item_after(0)
    assert package.last_change.operation == "insert_item_after"
    assert any("w15:repeatingSectionItem" in w for w in warnings_of(package))

    part = package.custom_xml_parts.get_item("{5D7BA57F-1E52-4637-9F82-2D4025768D4F}")
    package.changes.clear()
    assert section.xml_mapping.set_mapping("/invoice[1]/lines[1]/lineitem[1]", "", part)
    assert any("w15:dataBinding" in w for w in warnings_of(package))


def test_nothing_warns_on_a_document_that_is_already_mode_15():
    package = sample(WORD_2013)

    package.changes.clear()
    control = package.body.paragraphs[0].insert_content_control("RepeatingSection")
    assert package.last_change.warnings == ()

    package.changes.clear()
    control.color = "#FF0000"
    assert package.last_change.warnings == ()

    package.changes.clear()
    control.appearance = "Hidden"
    assert package.last_change.warnings == ()


def test_raising_the_mode_stops_the_warning():
    package = sample(WORD_2010)
    package.compatibility_mode = 15

    package.changes.clear()
    package.body.paragraphs[0].insert_content_control("RepeatingSection")

    assert package.last_change.warnings == ()
