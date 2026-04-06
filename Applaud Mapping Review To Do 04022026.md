# Applaud Mapping Review To Do — 04/02/2026

Audit of `fbdi_applaud_mapping.xlsx` (Sheet2). 183 total Applaud tables.

---

## Action Required: YES Status with No Mapping Value (5 rows)

These rows have `Status=YES` but the FBDI Template Mappings column (E) is blank. Either fill in the correct mapping or change status to UNMAPPED.

| Row | Applaud Table |
|-----|--------------|
| 153 | T_BANKS_BRANCHES |
| 154 | T_BPA_PO_HEADERS_INTERFACE |
| 155 | T_BPA_PO_LINES_INTERFACE |
| 156 | T_BPA_PO_LINE_LOCATIONS |
| 162 | T_DOO_ORDER_LINES_EFF_B |

---

## No Mapping — Known/Expected Gaps (41 rows)

These are not errors — they reflect tables with no FBDI counterpart in scope, or files that couldn't be read. Kept here for awareness.

### UNMAPPED (37) — No FBDI template exists for these in scope

- T_BANKS (17)
- T_CSE_ASSETS_INT (20)
- T_OPERATION_ITEMS (82)
- T_OPERATION_RESOURCES (83)
- T_SO_ATTACHMENTS (120)
- T_TRADING_PARTNERS_IMPORT (121)
- T_WORK_DEFINITION_AND_OPS (130)
- T_WSH_TRANSIT_TIMES (145)
- T_AP_INVOICE_ATTACHMENTS (146)
- T_AR_RECEIPTS (149)
- T_ASSIGNMENT (150)
- T_ASSIGNMENTSUPERVISOR (151)
- T_ASSIGNMENTWORKMEASURE (152)
- T_CONTRACT_HDR_FLEX_P (157)
- T_CONTRACT_HDR_P (158)
- T_CONTRACT_LINES_P (159)
- T_CONTRACT_PARTY_CONTACT_P (160)
- T_CONTRACT_PARTY_P (161)
- T_FA_DISPOSALS (163)
- T_INV_INVENTORY_LOCATOR (164)
- T_INV_INVENTORY_SUBINVENTORY (165)
- T_INV_ITEM_LOCATOR (166)
- T_INV_ITEM_SUBINVENTORY (167)
- T_INV_SERIAL_NUMBERS_INT (168)
- T_ORA_INV_SUBINVENTORY_LOCATOR (169)
- T_ORA_INV_SUBINV_SOURCE (170)
- T_ORA_WSH_TRANSIT_TIME_VALUES (171)
- T_PERSONADDRESS (172)
- T_PERSONEMAIL (173)
- T_PERSONLEGISLATIVEDATA (174)
- T_PERSONNAME (175)
- T_PERSONPHONE (176)
- T_PERSONUSERINFORMATION (177)
- T_PRICING_PROFILE (180)
- T_WORKER (182)
- T_WORKRELATIONSHIP (183)
- T_WORKTERMS (184)

### FILE_TOO_LARGE (2) — Skipped at comparison time

- T_AP_INVOICE_INT (147)
- T_AP_INVOICE_LINES (148)

### FILE_ERROR (2) — Unreadable xlsm files

- T_POZ_SUPPLIER_SITES_INT (178)
- T_POZ_SUP_THIRDPARTY_INT (179)

---

## FYI: Multi-Template Mappings (31 rows) — These Are Valid

Your initial assumption was "no Applaud table should map to more than 1 FBDI tab." This is not correct — Oracle intentionally reuses the same interface table across multiple FBDI templates (e.g., the same PO interface table appears in Blanket PA, Contract PA, and Purchase Order templates). These multi-mappings are by Oracle design. No changes needed.

Full list of the 31 multi-mapped tables for reference:

| Row | Applaud Table | Templates |
|-----|--------------|-----------|
| 9 | T_AWARD_KEYWORDS | ImportAwards; ImportGrantsPersonnel |
| 18 | T_BPA_PO_GA_ORG_ASSIGN_INTERFA | POBlanketPurchaseAgreementImportTemplate; POContractPurchaseAgreementImportTemplate |
| 35 | T_EGO_ITEM_INTF_EFF_B | ChangeOrderImportTemplate; ItemImportTemplate |
| 36 | T_EGP_COMPONENTS_INTERFACE | ChangeOrderImportTemplate; ItemStructureImportTemplate |
| 37 | T_EGP_ITEM_ATTACHMENTS_INTF | ChangeOrderImportTemplate; ItemImportTemplate |
| 39 | T_EGP_ITEM_RELATIONSHIPS_INTF | ChangeOrderImportTemplate; ItemImportTemplate |
| 41 | T_EGP_REF_DESGS_INTERFACE | ChangeOrderImportTemplate; ItemStructureImportTemplate |
| 42 | T_EGP_STRUCTURES_INTERFACE | ChangeOrderImportTemplate; ItemStructureImportTemplate |
| 43 | T_EGP_SUB_COMPS_INTERFACE | ChangeOrderImportTemplate; ItemStructureImportTemplate |
| 44 | T_EGP_SYSTEM_ITEMS_INTERFACE | ChangeOrderImportTemplate; ItemImportTemplate |
| 46 | T_FA_ADJUSTMENTS | FixedAssetMassAdjustmentsImportTemplate; FixedAssetMassRevaluationsImportTemplate |
| 70 | T_IBY_TEMP_EXT_BANK_ACCT | SupplierBankAccountImportTemplate; UploadCustomersTemplate |
| 73 | T_INV_LPN_INTERFACE | PerformShippingTransactionImportTemplate; ReceivingReceiptImportTemplate |
| 74 | T_INV_SERIAL_NUMBERS_INTERFACE | InterfacedPickTransactionsImportTemplate; InventoryTransactionImportTemplate; PerformShippingTransactionImportTemplate; ReceivingReceiptImportTemplate |
| 77 | T_INV_TRANSACT_LOTS_INTERFACE | InterfacedPickTransactionsImportTemplate; InventoryTransactionImportTemplate; PerformShippingTransactionImportTemplate; ReceivingReceiptImportTemplate |
| 85 | T_PJC_TXN_XFACE_STAGE_ALL | 6x ProjectUnprocessed*ExpenditureItemImportTemplate variants |
| 95 | T_PO_HEADERS_INTERFACE | POBlanketPurchaseAgreementImportTemplate; POContractPurchaseAgreementImportTemplate; POPurchaseOrderImportTemplate |
| 96 | T_PO_LINES_INTERFACE | POBlanketPurchaseAgreementImportTemplate; POPurchaseOrderImportTemplate |
| 97 | T_PO_LINE_LOCATIONS_INTERFACE | POBlanketPurchaseAgreementImportTemplate; POPurchaseOrderImportTemplate |
| 101 | T_PROJ_ENT_RES_INTERFACE | ProjectEnterpriseExpenseResourcesImportTemplate; ProjectEnterpriseResourcesImportTemplate |
| 103 | T_QP_MATRIX_DIMENSIONS_INT | DiscountListImportTemplate; PriceListsImportBatchTemplate |
| 104 | T_QP_MATRIX_RULES_INT | DiscountListImportTemplate; PriceListsImportBatchTemplate |
| 123 | T_WIE_INT_BATCHES_VL | MaintenanceWorkDefinitionTemplate; ProcessWorkDefinitionTemplate; WorkDefinitionTemplate |
| 124 | T_WIS_WD_DETAILS_INT-ARES | ProcessWorkDefinitionTemplate; WorkDefinitionTemplate |
| 132 | T_WO_BATCHES | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 133 | T_WO_HEADER | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 136 | T_WO_OPERATIONS | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 137 | T_WO_OPERATION_MATERIALS | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 138 | T_WO_OPERATION_RESOURCES | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 139 | T_WO_OPN_RES_INSTANCES | MaintenanceWorkOrderTemplate; ProcessWorkOrderTemplate; WorkOrderTemplate |
| 140 | T_WO_PRODUCT_LOT_NUMBERS | ProcessWorkOrderTemplate; WorkOrderTemplate |

---

## Summary

| Category | Count |
|----------|-------|
| YES but blank mapping — fix needed | 5 |
| UNMAPPED — no FBDI counterpart in scope | 37 |
| FILE_TOO_LARGE — skipped at comparison | 2 |
| FILE_ERROR — unreadable xlsm | 2 |
| Multi-template mappings — valid, no action | 31 |
| Clean single-template mappings | ~106 |
