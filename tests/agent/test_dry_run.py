"""``pkg.dry_run()``: edits on a copy, and the real document untouched.

CR-003 section 3.4, Phase D.
"""

from __future__ import annotations

import pytest
from conftest import WITHOUT_PARA_IDS, part_bytes, sample

from docx4j_py.model.content import ContentError
from docx4j_py.model.content.trial import TrialPackage


def test_a_trial_edits_a_copy_and_reports_what_it_would_do(report):
    before = report.body.text

    with report.dry_run() as trial:
        assert isinstance(trial, TrialPackage)
        count = trial.body.replace_text("lazy", "energetic")
        assert count == 2
        assert "energetic" in trial.body.text
        assert trial.last_change.operation == "replace_text"
        assert len(trial.changes) == 1

    assert report.body.text == before, "the real document never saw it"
    assert report.changes == [], "and its own reports are untouched"

    # having seen the number, commit
    assert report.body.replace_text("lazy", "energetic") == 2
    assert report.last_change.operation == "replace_text"


def test_the_real_package_is_byte_for_byte_what_it_was():
    package = sample(WITHOUT_PARA_IDS)
    package.body  # noqa: B018 - unmarshal the main part first, as a caller would
    before = package.save()

    with package.dry_run() as trial:
        trial.body.insert_paragraph("added in the trial")
        trial.body.paragraphs[0].delete()
        trial.body.clear()

    assert package.save() == before
    assert part_bytes(package.save()) == part_bytes(before)


def test_the_trial_sees_the_document_as_it_is(report):
    report.body.insert_paragraph("added before the trial")

    with report.dry_run() as trial:
        assert trial.body.paragraphs[-1].text == "added before the trial"
        assert trial.outline().stats.paragraphs == report.outline().stats.paragraphs
        assert trial.find("lazy")[0].ordinal == report.find("lazy")[0].ordinal
        assert trial.describe().page.orientation == "portrait"
        assert trial.paragraph_at("body/0").text == "Quarterly Report"


def test_a_trial_addresses_and_ids_match_the_commit_that_follows(report):
    with report.dry_run() as trial:
        made = trial.body.insert_paragraph("Appendix", style="Heading 1")
        trial_id = made.para_id
        trial_address = made.address

    committed = report.body.insert_paragraph("Appendix", style="Heading 1")
    assert committed.para_id == trial_id, "the trial's generator is seeded like the real one"
    assert committed.address == trial_address


def test_a_trial_covers_the_headers_too():
    package = sample("Headers.docx")
    header = next(p for p in package.bodies() if p.prefix.startswith("header:"))
    before = header.text

    with package.dry_run() as trial:
        trial_header = next(b for b in trial.bodies() if b.prefix == header.prefix)
        trial_header.insert_paragraph("trial header line")
        assert "trial header line" in trial_header.text

    assert header.text == before


def test_a_trial_cannot_be_saved(report):
    with report.dry_run() as trial, pytest.raises(ContentError) as raised:
        trial.save()
    assert raised.value.code == "dry_run.save"
    assert "real package" in raised.value.hint


def test_a_trial_of_a_trial(report):
    with report.dry_run() as trial, trial.dry_run() as inner:
        inner.body.insert_paragraph("two levels down")
        assert len(inner.body) == len(report.body) + 1
    assert len(trial.body) == len(report.body)
