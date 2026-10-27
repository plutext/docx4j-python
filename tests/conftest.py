"""The whole suite saves with the validator on (CR-020 Phase B, Q5).

Every ``package.save()`` a test makes, directly or through the content API, the
python-docx facade or a session, runs :func:`docx4j_py.validator.check_for_save`
first, as ``save(check=True)`` would: content this session introduced that
Word would refuse or read otherwise (misplaced content, a table cell that does
not end with a paragraph, an integer measurement past what Word stores)
raises :class:`~docx4j_py.model.content.errors.SaveCheckError`, so an engine
edit that makes one fails the test that made it. What a source file already
held is saved as it is (``last_save.findings``).

A test that saves a tree it built invalid on purpose opts out where it saves:
``save(check=False)``, or ``with checking(False):`` around a save it does not
make itself. ``tests/README.md`` lists them. ``DOCX4J_PY_SAVE_CHECK=0`` turns the
check off for a run, to measure what it costs.
"""

from __future__ import annotations

import os

import pytest

from docx4j_py.openpackaging.save import checking


@pytest.fixture(autouse=True, scope="session")
def _every_save_is_checked():
    """CR-020 Q5: ``save(check=None)`` means ``check=True`` for the whole run."""
    with checking(os.environ.get("DOCX4J_PY_SAVE_CHECK", "1") != "0"):
        yield
