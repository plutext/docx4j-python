"""Scripted sessions of tool-shaped calls, over the corpus.

CR-003 section 7: "scripted sessions of tool-shaped calls (``outline``,
``find``, edit by address, ``ChangeReport`` checked, ``dry_run`` then commit)
over the corpus". Every step here is one tool call: JSON in, JSON out, and the
only thing carried from one step to the next is a string.
"""

from __future__ import annotations

import json

from conftest import SEED, WITH_PARA_IDS, WITHOUT_PARA_IDS, reloaded, sample

from docx4j_py import DocumentSession, WordprocessingMLPackage


def tool(result: object) -> object:
    """What a server would return: the result through JSON and back."""
    return json.loads(json.dumps(result))


def test_a_session_over_a_document_that_has_no_para_ids():
    package = sample(WITHOUT_PARA_IDS)

    # 1. outline: what is in this document, and what can I address?
    outline = tool(package.outline(max_chars=40).to_dict())
    assert outline["stats"]["paragraphs"] >= 3
    first = outline["entries"][0]
    assert first["address"] == first["ordinal"] == "body/0"
    assert "para_id" not in first, "this document has none, so the ordinal is the handle"

    # 2. find: where does a word occur?
    hits = tool([hit.to_dict() for hit in package.find("document", limit=5)])
    assert hits, "the sample says document"
    address = hits[0]["address"]

    # 3. edit by address
    paragraph = package.paragraph_at(address)
    made = paragraph.insert_paragraph("Added by an agent.", location="After")

    # 4. the ChangeReport is the tool result
    change = tool(package.last_change.to_dict())
    assert change["operation"] == "insert_paragraph"
    assert change["addresses"] == [made.address]
    assert change["parts_touched"] == ["/word/document.xml"]
    assert change["text_after"] == "Added by an agent."
    assert change["moved"], "the blocks after it shifted, and the report says which"

    # 5. save, reload, read back
    again = reloaded(package)
    assert "Added by an agent." in again.body.text
    assert again.paragraph_at(made.ordinal).text == "Added by an agent."


def test_a_session_over_a_document_that_has_para_ids():
    package = sample(WITH_PARA_IDS)

    outline = package.outline()
    stamped = [entry for entry in outline.entries if entry.para_id]
    assert stamped, "DrawingML_GraphicData_wps.docx carries w14:paraId"
    address = stamped[-1].address
    assert address.startswith("w14:")

    # an insert in front of it: the paraId address still resolves, the ordinal
    # has moved, and the report says so (CR-003 section 3.4)
    before_ordinal = package.paragraph_at(address).ordinal
    package.body.insert_paragraph("Preamble", location="Start")

    after_ordinal = package.paragraph_at(address).ordinal
    assert after_ordinal != before_ordinal, "the ordinal moved"
    assert (before_ordinal, after_ordinal) in package.last_change.moved, "and the report says so"

    again = reloaded(package)
    assert again.paragraph_at(address) is not None, "the handle survives the round trip"


def test_dry_run_then_commit_through_a_session(tmp_path):
    path = tmp_path / "work.docx"
    path.write_bytes(sample(WITHOUT_PARA_IDS).save())

    with DocumentSession() as session:
        handle = session.open(path)

        with session.use(handle) as package:
            package.id_seed = SEED

            # preview
            with package.dry_run() as trial:
                count = trial.body.replace_text("document", "report")
                preview = tool(trial.last_change.to_dict())
            assert preview["operation"] == "replace_text"
            assert count >= 1

            # the real document is still what it was
            assert "report" not in package.body.text

            # commit, having seen the preview
            assert package.body.replace_text("document", "report") == count
            assert tool(package.last_change.to_dict())["text_after"] == "report"

        session.save(handle, overwrite=True)

    assert "report" in WordprocessingMLPackage.load(path).body.text


def test_the_same_seed_and_the_same_calls_give_the_same_bytes():
    def run() -> bytes:
        package = sample(WITHOUT_PARA_IDS)
        package.id_seed = SEED
        package.assigns_para_ids = True
        body = package.body
        body.ensure_para_ids()
        title = body.insert_paragraph("Report", location="Start", style="Heading 1")
        title.alignment = "Centered"
        body.paragraph_at(contains="document").insert_paragraph("Second", location="After")
        body.replace_text("document", "report")
        body.find("report")[0].range(body).font.italic = True
        return package.save()

    first, second = run(), run()
    assert first == second, "same document, same calls, same bytes (CR-003 section 3.4)"

    # and the reports agree too, bar their timestamps
    def reports() -> list[dict]:
        package = sample(WITHOUT_PARA_IDS)
        package.id_seed = SEED
        package.assigns_para_ids = True
        package.body.ensure_para_ids()
        package.body.insert_paragraph("Report", location="Start", style="Heading 1")
        return [
            {key: value for key, value in change.to_dict().items() if key != "at"}
            for change in package.changes
        ]

    assert reports() == reports()


def test_a_scripted_session_over_every_docx_in_the_corpus():
    names = [
        "2010-sample1.docx",
        "2010-glow-then-AlternateContent.docx",
        "DrawingML_GraphicData_wps.docx",
        "Headers.docx",
        "Images.docx",
        "invoice2013.docx",
        "sample-docx.docx",
        "Symbols.docx",
        "tables.docx",
        "toc.docx",
        "w14_texteffects.docx",
    ]
    for name in names:
        package = sample(name)

        outline = package.outline(max_chars=40)
        assert outline.stats.paragraphs > 0, name
        for address in outline.addresses():
            assert package.element_at(address) is not None, f"{name} {address}"

        description = package.describe()
        assert description.parts, name
        assert description.page.width_pt > 0, name

        first = outline.entries[0]
        package.paragraph_at(first.address) if first.kind == "paragraph" else None

        package.body.insert_paragraph("agent was here")
        change = package.last_change
        assert change.operation == "insert_paragraph"
        assert change.addresses

        again = reloaded(package)
        assert "agent was here" in again.body.text, name
