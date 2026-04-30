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
