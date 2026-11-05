"""Tests for tools/download_and_clear.py — static toc.htm discovery, the
module map, and the post-download "did the file vanish?" check.

No network: discover() takes an injected fetch(url) -> html.
"""

import importlib.util
import json
import sys
from pathlib import Path

import pytest

TOOL = Path(__file__).resolve().parent.parent / "tools" / "download_and_clear.py"


@pytest.fixture(scope="module")
def dac():
    spec = importlib.util.spec_from_file_location("download_and_clear", TOOL)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["download_and_clear"] = mod
    spec.loader.exec_module(mod)
    return mod


SCM = "https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/26d/oefsc/"
FIN = "https://docs.oracle.com/en/cloud/saas/financials/26d/oefbf/"
PPM = "https://docs.oracle.com/en/cloud/saas/project-management/26d/oefpp/"
PRC = "https://docs.oracle.com/en/cloud/saas/procurement/26d/oefbp/"
WEB = "https://www.oracle.com/webfolder/technetwork/docs/fbdi-26d/fbdi/xlsm/"


class TestParseToc:
    def test_relative_topics_resolved_deduped_fragments_dropped(self, dac):
        html = """
        <link href="/sp_common/book-template/css/book.css">
        <a href="index.html">Home</a>
        <a href="catalog-3181.html#catalog-3181">Catalog</a>
        <a href="catalog-3181.html#catalog-3181">Catalog again</a>
        <a href="requirementimport-3170.html#requirementimport-3170">Requirements</a>
        <a href="http://purl.org/dc/terms/">dc</a>
        <a href="https://docs.oracle.com/en/cloud/saas/financials/26d/oefbf/x.html">other book</a>
        """
        assert dac.parse_toc(html, SCM + "toc.htm") == [
            SCM + "catalog-3181.html",
            SCM + "requirementimport-3170.html",
        ]

    def test_empty_toc(self, dac):
        assert dac.parse_toc("<html></html>", SCM + "toc.htm") == []


class TestExtractXlsmLinks:
    def test_absolute_and_relative_xlsm_hrefs(self, dac):
        html = f"""
        <a href="{WEB}RequirementImportTemplate.xlsm">template</a>
        <a href="{WEB}RequirementImportTemplate.xlsm">dup</a>
        <a href="files/Local.XLSM">relative, upper-case ext</a>
        <a href="guide.pdf">not a template</a>
        """
        assert dac.extract_xlsm_links(html, SCM + "requirementimport-3170.html") == [
            WEB + "RequirementImportTemplate.xlsm",
            SCM + "files/Local.XLSM",
        ]

    def test_filename_is_unquoted_basename(self, dac):
        assert dac.xlsm_filename(WEB + "My%20Template.xlsm?x=1") == "My Template.xlsm"


def _site(pages: dict[str, str]):
    """fetch() over a dict; unknown URLs raise like a 404 would."""
    def fetch(url):
        if url not in pages:
            raise RuntimeError(f"404 {url}")
        return pages[url]
    return fetch


def _toc(*topics):
    return "".join(f'<a href="{t}#x">{t}</a>' for t in topics)


class TestDiscover:
    def _pages(self):
        return {
            PPM + "toc.htm": _toc("projects-1.html"),
            PPM + "projects-1.html": f'<a href="{WEB}ProjectImportTemplate.xlsm">x</a>',
            FIN + "toc.htm": _toc("payables-1.html"),
            FIN + "payables-1.html": f'<a href="{WEB}ApInvoicesImportTemplate.xlsm">x</a>',
            PRC + "toc.htm": _toc("changeorders-1.html"),
            PRC + "changeorders-1.html": f'<a href="{WEB}ChangeOrderImportTemplate.xlsm">x</a>',
            # SCM links ChangeOrder too; the earlier book (Procurement) wins.
            SCM + "toc.htm": _toc("changeorders-3180.html", "requirementimport-3170.html"),
            SCM + "changeorders-3180.html": f'<a href="{WEB}ChangeOrderImportTemplate.xlsm">x</a>',
            SCM + "requirementimport-3170.html": f'<a href="{WEB}RequirementImportTemplate.xlsm">x</a>',
        }

    def test_finds_every_linked_template_with_its_book_module(self, dac):
        disc = dac.discover("26D", _site(self._pages()), workers=2)
        assert disc.errors == []
        assert disc.pages == 5
        assert {n: m for n, (_u, m) in disc.files.items()} == {
            "ProjectImportTemplate.xlsm": "Project Management",
            "ApInvoicesImportTemplate.xlsm": "Financials",
            "ChangeOrderImportTemplate.xlsm": "Procurement",
            "RequirementImportTemplate.xlsm": "Supply Chain & Manufacturing",
        }
        assert disc.files["RequirementImportTemplate.xlsm"][0] == WEB + "RequirementImportTemplate.xlsm"

    def test_version_token_is_lowercased(self, dac):
        disc = dac.discover("26d", _site(self._pages()), workers=2)
        assert len(disc.files) == 4

    def test_failed_topic_page_is_reported_not_swallowed(self, dac):
        pages = self._pages()
        del pages[SCM + "requirementimport-3170.html"]
        disc = dac.discover("26D", _site(pages), workers=2)
        assert "RequirementImportTemplate.xlsm" not in disc.files
        assert len(disc.errors) == 1
        assert "requirementimport-3170.html" in disc.errors[0]

    def test_failed_or_empty_toc_is_reported(self, dac):
        pages = self._pages()
        del pages[FIN + "toc.htm"]
        pages[PRC + "toc.htm"] = "<html>no topics</html>"
        disc = dac.discover("26D", _site(pages), workers=2)
        assert any(FIN + "toc.htm" in e for e in disc.errors)
        assert any(PRC + "toc.htm" in e and "no topic pages" in e for e in disc.errors)
        # Procurement's TOC is empty, so SCM now supplies ChangeOrder.
        assert disc.files["ChangeOrderImportTemplate.xlsm"][1] == "Supply Chain & Manufacturing"


class TestFindMissing:
    def test_absent_and_empty_files_are_missing(self, dac, tmp_path):
        (tmp_path / "Kept.xlsm").write_bytes(b"PK")
        (tmp_path / "Empty.xlsm").write_bytes(b"")
        assert dac.find_missing(
            ["Kept.xlsm", "Empty.xlsm", "Quarantined.xlsm"], str(tmp_path)
        ) == ["Empty.xlsm", "Quarantined.xlsm"]


class TestWriteModuleMap:
    def test_writes_upper_case_release_dir_and_adds_manual_file(self, dac, tmp_path):
        out = dac.write_module_map({"A.xlsm": "Financials"}, "26d", str(tmp_path))
        assert Path(out) == tmp_path / "26D" / "file_modules.json"
        assert json.loads(Path(out).read_text(encoding="utf-8")) == {
            "A.xlsm": "Financials",
            "RapidImplementationForCashManagement.xlsm": "Financials",
        }


class TestModuleFromBaseUrl:
    def test_toc_url_maps_to_module(self, dac):
        assert dac.module_from_base_url(FIN + "toc.htm") == "Financials"

    def test_unknown_slug_raises(self, dac):
        with pytest.raises(ValueError):
            dac.module_from_base_url("https://docs.oracle.com/en/cloud/saas/hcm/26d/x/toc.htm")


class TestDownloadAll:
    class _Resp:
        def __init__(self, body, fail=False):
            self.body, self.fail = body, fail
        def __enter__(self):
            return self
        def __exit__(self, *a):
            return False
        def raise_for_status(self):
            if self.fail:
                import requests
                raise requests.exceptions.HTTPError("404")
        def iter_content(self, chunk_size):
            yield self.body

    class _Session:
        def __init__(self, responses):
            self.responses = responses
        def get(self, url, stream=True, timeout=None):
            return self.responses[url]

    def test_writes_files_and_reports_failures_without_partials(self, dac, tmp_path):
        files = {
            "Good.xlsm": (WEB + "Good.xlsm", "Financials"),
            "Bad.xlsm": (WEB + "Bad.xlsm", "Financials"),
        }
        session = self._Session({
            WEB + "Good.xlsm": self._Resp(b"PK\x03\x04"),
            WEB + "Bad.xlsm": self._Resp(b"", fail=True),
        })
        failures = dac.download_all(files, str(tmp_path), session)
        assert [n for n, _ in failures] == ["Bad.xlsm"]
        assert (tmp_path / "Good.xlsm").read_bytes() == b"PK\x03\x04"
        assert sorted(p.name for p in tmp_path.iterdir()) == ["Good.xlsm"]
