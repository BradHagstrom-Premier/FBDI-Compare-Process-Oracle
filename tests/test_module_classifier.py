"""Tests for module_from_base_url — Oracle docs URL → canonical module name."""

import pytest

from tools.download_and_clear import module_from_base_url


class TestModuleFromBaseUrl:
    def test_financials_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/financials/26b/oefbf/index.html"
        assert module_from_base_url(url) == "Financials"

    def test_procurement_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/procurement/26b/oefbp/index.html"
        assert module_from_base_url(url) == "Procurement"

    def test_supply_chain_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/supply-chain-and-manufacturing/26b/oefsc/index.html"
        assert module_from_base_url(url) == "Supply Chain & Manufacturing"

    def test_project_management_url(self):
        url = "https://docs.oracle.com/en/cloud/saas/project-management/26b/oefpp/index.html"
        assert module_from_base_url(url) == "Project Management"

    def test_unknown_url_raises(self):
        with pytest.raises(ValueError, match="Unknown Oracle module URL"):
            module_from_base_url("https://docs.oracle.com/en/cloud/saas/hcm/26b/x/y.html")


def test_write_module_map_round_trip(tmp_path, monkeypatch):
    """write_module_map produces well-formed JSON with FSM file added."""
    from tools.download_and_clear import write_module_map
    import json

    file_modules = {
        "AutoInvoiceImportTemplate.xlsm": "Financials",
        "ItemImportTemplate.xlsm": "Supply Chain & Manufacturing",
    }
    out_path = write_module_map(file_modules, "26C", str(tmp_path))

    with open(out_path) as f:
        data = json.load(f)

    # FSM file auto-added
    assert data["RapidImplementationForCashManagement.xlsm"] == "Financials"
    # User-supplied entries preserved
    assert data["AutoInvoiceImportTemplate.xlsm"] == "Financials"
    assert data["ItemImportTemplate.xlsm"] == "Supply Chain & Manufacturing"
    # Sorted keys (deterministic output)
    keys = list(data.keys())
    assert keys == sorted(keys)
