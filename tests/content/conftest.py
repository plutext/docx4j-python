"""Fixtures for the content API's tests. CR-003 Phase B, section 7.

Two helpers carry most of the weight. :func:`reloaded` is the rule CR-003
section 7 sets for every phase --- "every mutation followed by save, reload and
read back" --- and :func:`part_bytes` is the other one, "untouched parts still
byte-identical through the content API".
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from docx4j_py import WordprocessingMLPackage, create_package, load

#: The sample documents this phase reads.
SAMPLES = ROOT / "samples"

#: The documents copied from docx4j that a single phase needs.
FIXTURES = ROOT / "tests" / "fixtures"


@pytest.fixture
def new_package() -> WordprocessingMLPackage:
    """A fresh document, as ``create_package()`` makes it."""
    package = create_package()
    package.id_seed = 20260916  # CR-003 section 3.4: the same calls, the same bytes
    return package


@pytest.fixture
def body(new_package: WordprocessingMLPackage):
    """The new document's body."""
    return new_package.body


def sample(name: str) -> WordprocessingMLPackage:
    """Load one of ``samples/``."""
    return load(SAMPLES / name)


def reloaded(package: WordprocessingMLPackage) -> WordprocessingMLPackage:
    """Save the package to bytes and load it again: what Word would see."""
    return WordprocessingMLPackage.load(package.save())


def part_bytes(data: bytes, *, relationships: bool = False) -> dict[str, bytes]:
    """Every entry of a saved package, by name.

    The relationship parts are left out by default: CR-002's own round-trip
    test makes the byte-for-byte promise for "every non-relationship entry",
    because a relationship part is rebuilt from the model on every save (its
    attribute order is the serialiser's, not the source's).
    """
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {
            entry.filename: archive.read(entry.filename)
            for entry in archive.infolist()
            if relationships or not entry.filename.endswith(".rels")
        }


def fixture(name: str) -> WordprocessingMLPackage:
    """Load one of ``tests/fixtures/``."""
    return load(FIXTURES / name)


# ---------------------------------------------------------------------------
# a comment thread, built by hand (CR-003 Phase G, section 15)
# ---------------------------------------------------------------------------
#
# No document in docx4j, in docx4j-core-ts or in this repository carries a
# *reply* thread: `tests/fixtures/comments-modern.docx` is Word-written and has
# all five parts, but one comment and no reply. Rather than fabricate a
# "Word-written" fixture, the thread below is built here, in the open, from XML
# written to match what Word writes --- two comments on one range, the second a
# reply to the first through `w15:paraIdParent`, and a third, resolved, on
# another paragraph.

_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
_W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
_W16CID = 'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"'
_W16CEX = 'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"'

_REF = '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{0}"/></w:r>'

THREAD_DOCUMENT = f"""<w:document {_W} {_W14}><w:body>
<w:p w14:paraId="0A0A0A0A"><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>
<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>
<w:r><w:t>beta</w:t></w:r>
<w:commentRangeEnd w:id="1"/><w:commentRangeEnd w:id="0"/>
{_REF.format(0)}{_REF.format(1)}
<w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>
<w:p w14:paraId="0B0B0B0B"><w:r><w:t>Delta</w:t></w:r>
<w:commentRangeStart w:id="2"/><w:r><w:t xml:space="preserve"> epsilon</w:t></w:r>
<w:commentRangeEnd w:id="2"/>{_REF.format(2)}</w:p>
</w:body></w:document>"""

_COMMENT = """<w:comment w:id="{id}" w:author="{author}" w:initials="{initials}" \
w:date="2026-09-17T09:0{id}:00Z"><w:p w14:paraId="{para_id}"><w:pPr>\
<w:pStyle w:val="CommentText"/></w:pPr>\
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>\
<w:r><w:t>{text}</w:t></w:r></w:p></w:comment>"""

THREAD_COMMENTS = f"""<w:comments {_W} {_W14}>
{_COMMENT.format(id=0, author="Ada Lovelace", initials="AL", para_id="AAAA0000", text="Is this right?")}
{_COMMENT.format(id=1, author="Bob Bones", initials="BB", para_id="BBBB0000", text="Yes, checked.")}
{_COMMENT.format(id=2, author="Ada Lovelace", initials="AL", para_id="CCCC0000", text="Done with this one.")}
</w:comments>"""

THREAD_COMMENTS_EX = f"""<w15:commentsEx {_W15}>
<w15:commentEx w15:paraId="AAAA0000" w15:done="0"/>
<w15:commentEx w15:paraId="BBBB0000" w15:paraIdParent="AAAA0000" w15:done="0"/>
<w15:commentEx w15:paraId="CCCC0000" w15:done="1"/>
</w15:commentsEx>"""

THREAD_PEOPLE = f"""<w15:people {_W15}>
<w15:person w15:author="Ada Lovelace"><w15:presenceInfo w15:providerId="AD" \
w15:userId="S::ada@example.com::11111111-2222-3333-4444-555555555555"/></w15:person>
<w15:person w15:author="Bob Bones"><w15:presenceInfo w15:providerId="None" \
w15:userId="bob@example.com"/></w15:person>
</w15:people>"""

THREAD_IDS = f"""<w16cid:commentsIds {_W16CID}>
<w16cid:commentId w16cid:paraId="AAAA0000" w16cid:durableId="1A1A1A1A"/>
<w16cid:commentId w16cid:paraId="BBBB0000" w16cid:durableId="1B1B1B1B"/>
<w16cid:commentId w16cid:paraId="CCCC0000" w16cid:durableId="1C1C1C1C"/>
</w16cid:commentsIds>"""

THREAD_EXTENSIBLE = f"""<w16cex:commentsExtensible {_W16CEX}>
<w16cex:commentExtensible w16cex:durableId="1A1A1A1A" w16cex:dateUtc="2026-09-17T09:00:00Z"/>
<w16cex:commentExtensible w16cex:durableId="1B1B1B1B" w16cex:dateUtc="2026-09-17T09:01:00Z"/>
<w16cex:commentExtensible w16cex:durableId="1C1C1C1C" w16cex:dateUtc="2026-09-17T09:02:00Z"/>
</w16cex:commentsExtensible>"""


def threaded_package() -> WordprocessingMLPackage:
    """A document with a comment thread and all five comment parts, built here.

    Two comments on ``beta`` in the first paragraph --- the second a **reply**
    to the first --- and a third, **resolved**, on ``epsilon`` in the second.
    The authors carry emails in ``w:people``, one in Word's Active Directory
    ``S::address::guid`` form and one plain.
    """
    from docx4j_py.openpackaging.content_types import ContentTypes
    from docx4j_py.openpackaging.parts.default_xml_part import DefaultXmlPart
    from docx4j_py.openpackaging.parts.namespaces import Namespaces
    from docx4j_py.openpackaging.parts.wml import (
        CommentsExtendedPart,
        CommentsPart,
        PeoplePart,
    )

    package = create_package()
    package.id_seed = 20260917
    main = package.main_document_part
    main.set_xml(THREAD_DOCUMENT)
    for factory, xml in (
        (CommentsPart, THREAD_COMMENTS),
        (CommentsExtendedPart, THREAD_COMMENTS_EX),
        (PeoplePart, THREAD_PEOPLE),
    ):
        part = factory()
        part.set_xml(xml)
        main.add_target_part(part)
    for name, content_type, relationship_type, xml in (
        ("/word/commentsIds.xml", ContentTypes.WORDPROCESSINGML_COMMENTS_IDS,
         Namespaces.COMMENTS_IDS, THREAD_IDS),
        ("/word/commentsExtensible.xml", ContentTypes.WORDPROCESSINGML_COMMENTS_EXTENSIBLE,
         Namespaces.COMMENTS_EXTENSIBLE, THREAD_EXTENSIBLE),
    ):
        part = DefaultXmlPart(name, content_type, relationship_type)
        part.set_xml(xml)
        main.add_target_part(part)
    return package
