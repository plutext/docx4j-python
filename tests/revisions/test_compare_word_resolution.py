"""``scripts/compare_word_resolution.py``, on the agreement fixtures. CR-010 section 12.

The script puts Word's own Accept All and Reject All saves of acceptance
artefact 8 against the copies ``scripts/acceptance.py`` resolves with the engine.
Word's saves of artefact 8 do not exist until the Word check makes them, so it
is held here where Word's saves already exist: each agreement fixture resolved
by ``revisions.accept_all`` / ``reject_all`` and saved, against Word's save of
the same resolution (``test_word_saves.py``'s cases), and artefact 8's file
names laid out with the probe's.
"""

from __future__ import annotations

import shutil

import compare_word_resolution
import pytest
from revision_helpers import FIXTURES, fixture, revisions
from test_word_saves import CASES, _id


def _resolved(tracked, accept, target):
    package = fixture(tracked)
    (revisions.accept_all if accept else revisions.reject_all)(package)
    package.save(target)
    return target


@pytest.mark.parametrize("case", CASES, ids=[_id(c) for c in CASES])
def test_ours_and_words_saves_compare_alike(case, tmp_path):
    tracked, golden, accept = case
    ours = _resolved(tracked, accept, tmp_path / "ours.docx")
    assert compare_word_resolution.compare(ours, FIXTURES / golden) == (0, "")


def _artefact_8_layout(tmp_path, word_accepted, word_rejected):
    ours, word = tmp_path / "acceptance", tmp_path / "acceptance-word"
    ours.mkdir()
    word.mkdir()
    _resolved("cr043/cr043-tracked.docx", True, ours / "8-tracked-changes-accepted.docx")
    _resolved("cr043/cr043-tracked.docx", False, ours / "8-tracked-changes-rejected.docx")
    shutil.copy(FIXTURES / word_accepted, word / "8-word-accepted.docx")
    shutil.copy(FIXTURES / word_rejected, word / "8-word-rejected.docx")
    return ["--ours", str(ours), "--word", str(word)]


def test_artefact_8s_pairs_by_default(tmp_path, capsys):
    args = _artefact_8_layout(tmp_path, "cr043/cr043-accept-all.docx", "cr043/cr043-reject-all.docx")
    assert compare_word_resolution.main(args) == 0
    out = capsys.readouterr().out
    assert "8-tracked-changes-accepted.docx against Word's 8-word-accepted.docx: 0 lines differ" in out
    assert "8-tracked-changes-rejected.docx against Word's 8-word-rejected.docx: 0 lines differ" in out
    assert "2 compared, 0 differ" in out


def test_a_difference_is_counted_and_shown(tmp_path, capsys):
    """Word's Reject All put against our Accept All: the lines differ, and the diff says which."""
    args = _artefact_8_layout(tmp_path, "cr043/cr043-reject-all.docx", "cr043/cr043-accept-all.docx")
    assert compare_word_resolution.main(args) == 1
    out = capsys.readouterr().out
    assert "2 compared, 2 differ" in out
    assert "--- Word: 8-word-accepted.docx" in out and "+++ ours: 8-tracked-changes-accepted.docx" in out

    count, diff = compare_word_resolution.compare(
        tmp_path / "acceptance" / "8-tracked-changes-accepted.docx", FIXTURES / "cr043/cr043-reject-all.docx"
    )
    assert count > 0 and diff


def test_pairs_given_and_a_missing_save(tmp_path, capsys):
    ours = _resolved("core-ts/check33/33a-deleted.docx", True, tmp_path / "33a.docx")
    golden = FIXTURES / "core-ts/check33/33b-accepted.docx"
    assert compare_word_resolution.main([str(ours), str(golden)]) == 0
    assert compare_word_resolution.main(["--ours", str(tmp_path), "--word", str(tmp_path / "nowhere")]) == 2
    assert "missing:" in capsys.readouterr().out
