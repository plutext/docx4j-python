"""``create_package`` writes real document properties, not empty parts.

Found by the Word acceptance run of 2026-09-12: the created documents showed no
``Application`` / ``AppVersion`` and an empty ``core.xml``. docx4j writes these only
when ``docx4j.App.write`` / ``docx4j.dc.write`` are set; here they are always
written for a new document (CR-002 section 12.8).
"""

from __future__ import annotations

import datetime
import io
import re
import zipfile

from docx4j_py import WordprocessingMLPackage, load
from docx4j_py.openpackaging.packages.wordprocessingml_package import (
    APPLICATION_NAME,
    app_version,
    default_core_properties,
)


def test_app_version_is_in_words_xx_yyyy_form():
    assert re.fullmatch(r"\d+\.\d{4}", app_version())


def test_created_package_has_application_and_app_version():
    pkg = WordprocessingMLPackage.create_package()
    app = pkg.get_part("/docProps/app.xml").contents
    assert app.application == APPLICATION_NAME
    assert app.app_version == app_version()


def test_created_package_has_created_and_modified_with_xsi_type():
    pkg = WordprocessingMLPackage.create_package()
    data = pkg.save()
    core_xml = zipfile.ZipFile(io.BytesIO(data)).read("docProps/core.xml").decode()
    assert '<dcterms:created xsi:type="dcterms:W3CDTF">' in core_xml
    assert '<dcterms:modified xsi:type="dcterms:W3CDTF">' in core_xml
    app_xml = zipfile.ZipFile(io.BytesIO(data)).read("docProps/app.xml").decode()
    # the default namespace, as Word writes it, not a prefixed root
    # Word's declaration line ends with CRLF, as the engine writes it
    assert app_xml.startswith(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    )
    assert "<Application>docx4j-python</Application>" in app_xml


def test_default_core_properties_use_utc_seconds():
    now = datetime.datetime(2026, 9, 12, 10, 30, 45, 123456, tzinfo=datetime.timezone(datetime.timedelta(hours=10)))
    core = default_core_properties(now)
    assert str(core.created.value) == "2026-09-12T00:30:45Z"
    assert str(core.modified.value) == "2026-09-12T00:30:45Z"


def test_created_properties_survive_a_reload():
    pkg = WordprocessingMLPackage.create_package()
    with load(pkg.save()) as again:
        app = again.get_part("/docProps/app.xml").contents
        core = again.get_part("/docProps/core.xml").contents
        assert app.application == APPLICATION_NAME
        assert core.created.value.value.year >= 2026  # DerivedElement -> W3CDTF -> XmlDateTime
        assert again.skipped == []
