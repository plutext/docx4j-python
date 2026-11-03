"""Revision and comment dates under ``markup_detail: author_date``: docx4j ``MarkupDatesTest``. CR-027 Phase B.

docx4j's ``MarkupDatesTest`` at ``docx4j-17.3.5`` (``fc71f12e6``; WD 0.46
REQ-702 and REQ-706): a revision's and a comment's date is the author's local
time from ``w:date``, its ``Z`` not written, with the UTC offset where
``dateUtc`` gives it; ``dateUtc`` with ``Z`` where it is alone; the author
alone where there is neither.

- ``measuredCases``: Word's four measured cases (``manual-redline-dates.docx``,
  ``tests/fixtures/revisions/cr068/``) in plain text, in Markdown and as each
  run's :attr:`TextRevision.date`;
- ``noDateAndComments``: a mark with no date as ``{--A: gone--}``, and a
  comment's ``dateUtc`` read through the comments ids and extensible parts,
  the parts built as docx4j's test builds them.

Then this port's own (section 4.2): the comment's last paragraph carrying a
``w14:paraId``, which is not the content API's last paragraph with an entry; a
``dateUtc`` with a fraction; the block model's JSON; and a reading that
unmarshals neither side part (REQ-204).
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from mini_docx import NS, MiniDocx

from docx4j_py import load
from docx4j_py.model.text import MarkdownExporter, extract, markdown, plain_text

ROOT = Path(__file__).resolve().parent.parent.parent
DATES = ROOT / "tests" / "fixtures" / "revisions" / "cr068" / "manual-redline-dates.docx"

#: docx4j's ``AUTHOR_DATE``: ``TextOptions.of(TextView.MARKUP).markupDetail(AUTHOR_DATE)``.
AUTHOR_DATE = {"view": "markup", "markup_detail": "author_date"}

_COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
_COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_IDS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
_IDS_REL = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
_EXTENSIBLE_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
_EXTENSIBLE_REL = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
_CID = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
_CEX = 'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"'


def _anchored(comment_id: int, text: str) -> str:
    return (
        f'<w:p><w:commentRangeStart w:id="{comment_id}"/><w:r><w:t>{text}</w:t></w:r>'
        f'<w:commentRangeEnd w:id="{comment_id}"/><w:r><w:commentReference w:id="{comment_id}"/></w:r></w:p>'
    )


def _with_comments(body: str, comments: str, ids: str, extensible: str) -> MiniDocx:
    """docx4j's ``MiniDocx.body(...).part(comments).part(commentsIds).part(commentsExtensible)``."""
    return (
        MiniDocx.body(body)
        .part("comments.xml", _COMMENTS_CT, _COMMENTS_REL, "rIdC", f"<w:comments {NS}>{comments}</w:comments>")
        .part("commentsIds.xml", _IDS_CT, _IDS_REL, "rIdI", f"<w16cid:commentsIds {_CID}>{ids}</w16cid:commentsIds>")
        .part(
            "commentsExtensible.xml",
            _EXTENSIBLE_CT,
            _EXTENSIBLE_REL,
            "rIdE",
            f"<w16cex:commentsExtensible {_CEX}>{extensible}</w16cex:commentsExtensible>",
        )
    )


# ---------------------------------------------------------------------------
# MarkupDatesTest
# ---------------------------------------------------------------------------

EXPECTED = [
    "{--manual redline, 2026-10-31T00:00:00: deleted D1--}",
    "{--manual redline, 2026-10-31T00:00:00Z: deleted D2--}",
    "{--manual redline, 2026-10-31T00:00:00+00:00: deleted D3--}",
    "{--manual redline, 2026-10-31T11:00:00+11:00: deleted D4--}",
]


def test_measured_cases():
    """``measuredCases``: Word's four measured cases (CR-068 phase 0), in plain text, in Markdown and in the block model."""
    pkg = load(DATES)
    r = extract(pkg, **AUTHOR_DATE)
    text = plain_text(r)
    for e in EXPECTED:
        assert e in text, text
    md = MarkdownExporter().export(r, pkg)
    for e in EXPECTED:
        assert e in md, md
    found = [run.revision.date for b in r.blocks for run in b.runs if run.revision is not None]
    assert found == [
        "2026-10-31T00:00:00",
        "2026-10-31T00:00:00Z",
        "2026-10-31T00:00:00+00:00",
        "2026-10-31T11:00:00+11:00",
    ]


def _no_date_and_comments() -> MiniDocx:
    return _with_comments(
        '<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:del w:id="5" w:author="A"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>'
        + _anchored(0, "One")
        + _anchored(1, "Two"),
        '<w:comment w:id="0" w:author="R" w:date="2026-10-31T11:00:00Z"><w:p w14:paraId="1A2B3C4D"><w:r><w:t>with utc</w:t></w:r></w:p></w:comment>'
        '<w:comment w:id="1" w:author="R" w:date="2026-10-31T11:00:00Z"><w:p w14:paraId="5E6F7A8B"><w:r><w:t>without</w:t></w:r></w:p></w:comment>',
        '<w16cid:commentId w16cid:paraId="1a2b3c4d" w16cid:durableId="0ABC1234"/>',
        '<w16cex:commentExtensible w16cex:durableId="0abc1234" w16cex:dateUtc="2026-10-31T00:00:00Z"/>',
    )


def test_no_date_and_comments():
    """``noDateAndComments``: a mark with no date writes its author alone; a comment's ``dateUtc`` is read through the ids and extensible parts."""
    pkg = _no_date_and_comments().load()
    r = extract(pkg, **AUTHOR_DATE)
    text = plain_text(r)
    assert "Kept {--A: gone--}" in text, text
    assert "{>>R, 2026-10-31T11:00:00+11:00: with utc<<}" in text, text
    assert "{>>R, 2026-10-31T11:00:00: without<<}" in text, text
    assert r.comments["0"].date == "2026-10-31T11:00:00+11:00"
    assert r.comments["1"].date == "2026-10-31T11:00:00"
    md = MarkdownExporter().export(r, pkg)
    assert "{--A: gone--}" in md, md


# ---------------------------------------------------------------------------
# this port's own
# ---------------------------------------------------------------------------


def test_markdowns_comments_take_the_form_too():
    """The Markdown writer's comment and mark, as the plain text's (``markdown._detail``)."""
    pkg = _no_date_and_comments().load()
    md = markdown(extract(pkg, **AUTHOR_DATE), pkg)
    assert "{>>R, 2026-10-31T11:00:00+11:00: with utc<<}" in md, md
    assert "{>>R, 2026-10-31T11:00:00: without<<}" in md, md


def test_the_author_detail_is_unchanged():
    """``markup_detail: author`` writes ``author:`` whatever the date, as before."""
    text = plain_text(extract(_no_date_and_comments().load(), view="markup", markup_detail="author"))
    assert "Kept {--A: gone--}" in text and "{>>R: with utc<<}" in text and "{>>R: without<<}" in text


def test_a_comments_utc_date_is_its_last_paragraph_carrying_a_para_id():
    """docx4j's ``lastParaId``: the last paragraph with a ``w14:paraId``, entry or not (section 4.2).

    The content API's ``creation_date`` reads the last paragraph *with an
    entry*, so the two differ here: the first paragraph has the entry, the
    last its own paragraph id and none. The text model gives no offset; the
    content API, the instant, unchanged (Q2).
    """
    pkg = _with_comments(
        _anchored(0, "One"),
        '<w:comment w:id="0" w:author="R" w:date="2026-10-31T11:00:00Z">'
        '<w:p w14:paraId="1A2B3C4D"><w:r><w:t>first</w:t></w:r></w:p>'
        '<w:p w14:paraId="0F0F0F0F"><w:r><w:t>last</w:t></w:r></w:p><w:p><w:r><w:t>no id</w:t></w:r></w:p></w:comment>',
        '<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="0ABC1234"/>',
        '<w16cex:commentExtensible w16cex:durableId="0ABC1234" w16cex:dateUtc="2026-10-31T00:00:00Z"/>',
    ).load()
    assert extract(pkg, **AUTHOR_DATE).comments["0"].date == "2026-10-31T11:00:00"
    (comment,) = pkg.body.get_comments()
    assert comment.creation_date == datetime.datetime(2026, 10, 31, tzinfo=datetime.UTC)


def test_a_utc_date_with_a_fraction_still_gives_the_offset():
    """``dates_utc`` keeps microseconds; the text model reads the date to the second, as ``RevisionDates`` does (section 4.2)."""
    pkg = _with_comments(
        _anchored(0, "One"),
        '<w:comment w:id="0" w:author="R" w:date="2026-10-31T11:00:00Z"><w:p w14:paraId="1A2B3C4D"><w:r><w:t>c</w:t></w:r></w:p></w:comment>',
        '<w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="0ABC1234"/>',
        '<w16cex:commentExtensible w16cex:durableId="0ABC1234" w16cex:dateUtc="2026-10-31T00:00:00.5Z"/>',
    ).load()
    assert extract(pkg, **AUTHOR_DATE).comments["0"].date == "2026-10-31T11:00:00+11:00"


def test_the_block_models_json_carries_the_form():
    """Q2: the form in every view's block model, not only under ``author_date``."""
    pkg = load(DATES)
    data = json.loads(json.dumps(extract(pkg, view="markup").to_dict()))
    found = [run["revision"]["date"] for b in data["blocks"] for run in b["runs"] if run.get("revision")]
    assert found == [
        "2026-10-31T00:00:00",
        "2026-10-31T00:00:00Z",
        "2026-10-31T00:00:00+00:00",
        "2026-10-31T11:00:00+11:00",
    ]


def test_a_reading_parses_neither_side_part():
    """REQ-204: the ids and extensible parts are read aside, so the package is saved as it was loaded."""
    data = _no_date_and_comments().save()
    pkg = load(data)
    main = pkg.main_document_part
    extract(pkg, **AUTHOR_DATE)
    assert not main.comments_ids_part.is_parsed and not main.comments_extensible_part.is_parsed
    assert not main.comments_part.is_unmarshalled
