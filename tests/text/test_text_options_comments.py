"""CR-035: one setting for comments. ``comments: true`` alone turns a comment story on, and ``comments`` leaves ``include``.

    .venv-fork/bin/python -m pytest tests/text/test_text_options_comments.py -q

The oracle is WD 0.91's section 10, REQ-501, REQ-504 and REQ-1002, not docx4j's
``WORD_TEXT_VALUES``, which keeps ``include: comments`` until docx4j makes the same change.
"""

from __future__ import annotations

import warnings
from typing import Any

import pytest

from docx4j_py.model.text import TextOptions
from docx4j_py.model.text.options import TextOptionError

OLD_SPELLING = "comments=True"


def old(make: Any) -> TextOptions:
    """`make()`, which must warn, naming the new form."""
    with pytest.warns(DeprecationWarning, match=OLD_SPELLING):
        return make()


# ---------------------------------------------------------------------------
# section 1's table
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("options", "reads"),
    [
        ({"profile": "word-text"}, True),
        ({"profile": "word-text", "comments": False}, False),
        ({"profile": "word-text", "include": ["headers", "footers"]}, True),
        ({"profile": "word-text", "comments": False, "include": ["headers", "footers"]}, False),
    ],
)
def test_section_1s_table(options: dict[str, Any], reads: bool) -> None:
    assert TextOptions(**options).comments is reads


def test_the_preset_is_the_wds() -> None:
    """REQ-1002: ``include: [headers, footers]``, ``comments: true``."""
    options = TextOptions(profile="word-text").to_dict()
    assert options["include"] == ["headers", "footers"]
    assert options["comments"] is True


def test_comments_true_alone_turns_the_story_on_with_no_warning() -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        options = TextOptions(comments=True)
    assert options.comments and "comments" in options.includes


def test_include_no_longer_speaks_for_comments() -> None:
    assert not TextOptions(include=["headers"]).comments


def test_includes_still_answers_which_stories_are_in_scope() -> None:
    assert TextOptions(comments=True, include=["headers"]).includes == {"comments", "headers"}
    assert TextOptions(include=["headers"]).includes == {"headers"}


# ---------------------------------------------------------------------------
# the old spelling, in every form
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "make",
    [
        lambda: TextOptions(include=["comments"]),
        lambda: TextOptions(include="comments"),
        lambda: TextOptions(include="headers, comments"),
        lambda: TextOptions.of({"include": ["comments"]}),
        lambda: TextOptions.of(include=["comments"]),
        lambda: TextOptions().set("include", ["comments"]),
        lambda: TextOptions().include("comments"),
        lambda: TextOptions().include("headers", "comments"),
    ],
    ids=["keyword", "string", "string of two", "mapping", "of's keyword", "set", "fluent", "fluent of two"],
)
def test_the_old_spelling_sets_comments_and_leaves_include(make: Any) -> None:
    options = old(make)
    assert options.comments
    assert "comments" not in options.to_dict()["include"]
    assert options.to_dict()["comments"] is True


def test_the_old_spelling_keeps_the_rest_of_the_list() -> None:
    assert old(lambda: TextOptions(include=["headers", "comments", "glossary"])).to_dict()["include"] == ["headers", "glossary"]


def test_the_warning_points_at_the_callers_line() -> None:
    with pytest.warns(DeprecationWarning) as caught:
        TextOptions(include=["comments"])
    assert caught[0].filename == __file__


def test_the_old_spelling_is_the_callers_against_a_preset() -> None:
    options = old(lambda: TextOptions(include=["comments"]))
    assert "comments" in options._explicit  # the one preset sets it on too, so only the record can show it
    options.set("profile", "word-text")
    assert options.comments


def test_the_old_spelling_is_an_include_of_the_callers_too() -> None:
    """``include=["comments"]`` didn't ask for headers, and a preset doesn't add them, as it didn't before."""
    options = old(lambda: TextOptions(include=["comments"], profile="word-text"))
    assert options.to_dict()["include"] == []


def test_an_option_it_doesnt_know_still_refuses() -> None:
    with pytest.raises(TextOptionError):
        TextOptions(include=["remarks"])


# ---------------------------------------------------------------------------
# the contradiction (Q4)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "make",
    [
        lambda: TextOptions(comments=False, include=["comments"]),
        lambda: TextOptions(include="headers comments", comments=False),
        lambda: TextOptions(comments="false", include=["comments"]),
        lambda: TextOptions.of({"comments": False, "include": ["comments"]}),
        lambda: TextOptions.of({"include": ["comments"]}, comments=False),
        lambda: TextOptions.of({" comments ": False, " include ": ["comments"]}),
    ],
    ids=["keywords", "string", "string false", "mapping", "mapping and keyword", "spaced names"],
)
def test_one_call_contradicting_itself_is_refused(make: Any) -> None:
    with pytest.raises(TextOptionError) as caught:
        make()
    assert caught.value.reason == "invalid" and caught.value.statement == "REQ-504"
    assert "take comments out of include" in str(caught.value)


def test_what_a_released_to_dict_writes_is_refused_naming_the_fix() -> None:
    """0.2.0's ``to_dict()`` after ``include=["comments"]``, which meant comments on (section 3.3)."""
    with pytest.raises(TextOptionError, match="saved by to_dict"):
        TextOptions.of({"comments": False, "include": ["comments"], "view": "accepted"})


def test_comments_true_beside_the_old_spelling_only_warns() -> None:
    assert old(lambda: TextOptions(comments=True, include=["comments"])).comments


def test_a_none_says_nothing() -> None:
    assert old(lambda: TextOptions.of({"comments": None, "include": ["comments"]})).comments
    assert not TextOptions.of({"comments": False, "include": None}).comments


def test_an_include_it_refuses_is_refused_for_that() -> None:
    with pytest.raises(TextOptionError, match="include"):
        TextOptions(comments=False, include=["remarks"])


def test_a_bad_comments_value_is_refused_for_that() -> None:
    with pytest.raises(TextOptionError, match="not true or false"):
        TextOptions(comments="maybe", include=["comments"])


@pytest.mark.parametrize(
    ("make", "reads"),
    [
        (lambda: TextOptions(comments=False).include("comments"), True),
        (lambda: TextOptions(comments=False).set("include", ["comments"]), True),
        (lambda: TextOptions(include=["comments"]).set("comments", False), False),
        (lambda: TextOptions().include("comments").set("comments", False), False),
    ],
)
def test_sequential_calls_arent_a_contradiction_the_last_stands(make: Any, reads: bool) -> None:
    assert old(make).comments is reads


def test_an_attribute_after_the_old_spelling_stands() -> None:
    options = old(lambda: TextOptions(include=["comments"]))
    options.comments = False
    assert not options.comments


# ---------------------------------------------------------------------------
# a preset, and the output
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "make",
    [
        lambda: TextOptions(comments=False, profile="word-text"),
        lambda: TextOptions(profile="word-text", comments=False),
        lambda: TextOptions(profile="word-text").set("comments", False),
        lambda: TextOptions(comments=False).set("profile", "word-text"),
    ],
    ids=["before", "after", "set after", "profile set after"],
)
def test_a_callers_comments_false_holds_under_the_preset(make: Any) -> None:
    assert not make().comments


def test_to_dict_round_trips() -> None:
    options = TextOptions(profile="word-text", comments=False, include=["glossary"])
    again = TextOptions.of({key: value for key, value in options.to_dict().items() if value is not None})
    assert again.to_dict() == options.to_dict()


def test_a_saved_mapping_in_the_old_form_loads() -> None:
    options = old(lambda: TextOptions.of({"comments": True, "include": ["headers", "comments"]}))
    assert options.comments and options.to_dict()["include"] == ["headers"]


def test_a_copy_keeps_the_one_setting() -> None:
    options = TextOptions(comments=True).copy()
    assert options.comments and options.to_dict()["comments"] is True
