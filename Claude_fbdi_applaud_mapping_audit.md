# FBDI ↔ Applaud Mapping Audit — 26B

**Generated:** 2026-04-24T15:18:58.130328+00:00
**Snapshot:** applaud_snapshot.json @ 2026-04-24T14:20:04.776814Z
**Catalog:** FBDI_Master_Catalog.xlsx 26B tab
**Prior mapping:** fbdi_applaud_mapping.xlsx

## Summary

Of 183 Applaud tables audited: 108 YES, 36 UNMAPPED, 35 NEEDS_REVIEW. 40 rows changed from prior.

## Needs Review (35 rows)

### T_AWARD_FED_DOM_ASSIST_PRG (prefix: TE9) — NEEDS_REVIEW
- **Prior:** YES → `ImportAwards / Award Assistance Listing Number`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ImportAwards / Awards` — name=NONE, keys=100%, cols=50% → L
  - `ImportAwards / Award Funding Sources` — name=NONE, keys=100%, cols=50% → L
  - `ImportAwards / Award Projects` — name=NONE, keys=100%, cols=50% → L
  - `ImportAwards / Award Project Funding Sources` — name=NONE, keys=100%, cols=50% → L
  - `ImportAwards / Award Keywords` — name=NONE, keys=100%, cols=50% → L

### T_AWARD_PRJ_TASK_BURDEN_SCHED (prefix: TF7) — NEEDS_REVIEW
- **Prior:** YES → `ImportAwards / Award Prj Task Burden Schedules`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ImportAwards / Award Prj Task Burden Schedules` — name=NONE, keys=100%, cols=80% → L
  - `ImportAwards / Award Projects` — name=NONE, keys=100%, cols=60% → L
  - `ProjectBudgetsImportTemplate / PJO_BUDGETS_XFACE` — name=NONE, keys=100%, cols=60% → L
  - `ImportAwards / Awards` — name=NONE, keys=100%, cols=40% → L
  - `ImportAwards / Award Project Funding Sources` — name=NONE, keys=100%, cols=40% → L

### T_BPA_PO_GA_ORG_ASSIGN_INTERFA (prefix: T76) — NEEDS_REVIEW
- **Prior:** YES → `POBlanketPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE; POContractPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `POBlanketPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE` — name=NONE, keys=67%, cols=33% → L
  - `POContractPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE` — name=NONE, keys=67%, cols=33% → L

### T_CSE_ASSETS_INT (prefix: T61) — NEEDS_REVIEW
- **Prior:** UNMAPPED → `CseInstalledBaseAssetImport / Assets`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `CseInstalledBaseAssetImport / Assets` — name=NONE, keys=67%, cols=89% → M
  - `MntMaintenanceProgramImport / Work Requirements` — name=NONE, keys=67%, cols=7% → L
  - `MaintenanceWorkOrderTemplate / Work order Asset` — name=NONE, keys=67%, cols=4% → L
  - `MntMaintenanceProgramImport / Affected Assets` — name=NONE, keys=67%, cols=4% → L
  - `CseMeterReadingsImport / MeterReadings` — name=NONE, keys=67%, cols=3% → L

### T_CSE_INT_BATCHES_B (prefix: T60) — NEEDS_REVIEW
- **Prior:** YES → `CseGenealogyBulkImport / Import Batches`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `CseInstalledBaseAssetImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `CseMeterReadingsImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `CseWarrantyCoverageImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `MaintenanceWorkOrderTemplate / Import Batches` — name=NONE, keys=100%, cols=100% → L
  - `MntMaintenanceProgramImport / Batches` — name=NONE, keys=100%, cols=100% → L

### T_GMS_PERSONNEL_INT (prefix: TF8) — NEEDS_REVIEW
- **Prior:** YES → `ImportGrantsPersonnel / Grants Personnel`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ProcessWorkOrderTemplate / Work Order Opn Res Instances` — name=NONE, keys=100%, cols=38% → L
  - `WorkOrderResourceTransactionTemplate / Resource Transaction Header` — name=NONE, keys=100%, cols=38% → L
  - `WorkOrderTemplate / Work Order Opn Res Instances` — name=NONE, keys=100%, cols=38% → L
  - `ProjectUnprocessedExpenseReportExpenditureItemImportTemplate / PJC_TXN_XFACE_STAGE_ALL` — name=NONE, keys=100%, cols=27% → L
  - `ProjectUnprocessedInventoryExpenditureItemImportTemplate / PJC_TXN_XFACE_STAGE_ALL` — name=NONE, keys=100%, cols=27% → L

### T_GMS_SPONSORS_INT (prefix: TG5) — NEEDS_REVIEW
- **Prior:** YES → `ImportFundingSources / Sponsors`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ImportFundingSources / Sponsor Account Details` — name=NONE, keys=50%, cols=33% → L
  - `ImportFundingSources / Sponsor References` — name=NONE, keys=50%, cols=22% → L
  - `ImportFundingSources / Sponsors` — name=NONE, keys=50%, cols=11% → L

### T_INV_TRANSACT_LOTS_INTERFACE (prefix: TK6) — NEEDS_REVIEW
- **Prior:** YES → `InterfacedPickTransactionsImportTemplate / INV_TRANSACTION_LOTS_INTERFACE; InventoryTransactionImportTemplate / INV_TRANSACTION_LOTS_INTERFACE; PerformShippingTransactionImportTemplate / INV_TRANSACTION_LOTS_INTERFACE; ReceivingReceiptImportTemplate / INV_TRANSACTION_LOTS_INTERFACE`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `InventoryTransactionImportTemplate / INV_TRANSACTIONS_INTERFACE` — name=NONE, keys=0%, cols=44% → L

### T_OPERATION_ITEMS (prefix: T95) — NEEDS_REVIEW
- **Prior:** UNMAPPED → `ProcessWorkDefinitionTemplate / Operation Items`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `ProcessWorkDefinitionTemplate / Operation Items` — name=EXACT, keys=0%, cols=8% → M
  - `CreateBillingEventsTemplate / Project Billing Events` — name=NONE, keys=100%, cols=8% → L
  - `ConfiguratorRedwoodRuleConversionTemplate / CZ_RW_RULE_CONVERSION` — name=NONE, keys=100%, cols=4% → L
  - `CustomerImportTemplate / HZ_IMP_PARTIES_T` — name=NONE, keys=100%, cols=4% → L
  - `ImportAwards / Award Organization Credits` — name=NONE, keys=100%, cols=4% → L

### T_OPERATION_RESOURCES (prefix: T96) — NEEDS_REVIEW
- **Prior:** UNMAPPED → `MaintenanceWorkOrderTemplate / Operation resources`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `MaintenanceWorkOrderTemplate / Operation resources` — name=EXACT, keys=0%, cols=8% → M
  - `ProcessWorkDefinitionTemplate / Operation Resources` — name=EXACT, keys=0%, cols=8% → M
  - `WorkDefinitionTemplate / Operation Resources` — name=EXACT, keys=0%, cols=8% → M
  - `ConfiguratorRedwoodRuleConversionTemplate / CZ_RW_RULE_CONVERSION` — name=NONE, keys=100%, cols=4% → L
  - `CreateBillingEventsTemplate / Project Billing Events` — name=NONE, keys=100%, cols=4% → L

### T_PJB_BILLING_EVENTS_INT (prefix: T38) — NEEDS_REVIEW
- **Prior:** YES → `CreateBillingEventsTemplate / Project Billing Events`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `CreateBillingEventsTemplate / Project Billing Events` — name=NONE, keys=100%, cols=17% → L

### T_POZ_SUP_ADDRESSES_INT (prefix: T08) — NEEDS_REVIEW
- **Prior:** YES → `SupplierAddressImportTemplate / POZ_SUPPLIER_ADDRESSES_INT`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ChangeOrderImportTemplate / EGO_CHANGES_INT` — name=NONE, keys=25%, cols=50% → M
  - `ChangeOrderImportTemplate / EGP_ITEM_ATTACHMENTS_INTF` — name=NONE, keys=25%, cols=47% → M
  - `EgpCatalogImportTemplate / Attachments` — name=NONE, keys=25%, cols=47% → M
  - `ItemImportTemplate / EGP_ITEM_ATTACHMENTS_INTF` — name=NONE, keys=25%, cols=47% → M
  - `ChangeOrderImportTemplate / EGP_SYSTEM_ITEMS_INTERFACE` — name=NONE, keys=25%, cols=43% → M

### T_POZ_SUP_CONTACT_ADDRESS_INT (prefix: T12) — NEEDS_REVIEW
- **Prior:** YES → `SupplierContactImportTemplate / POZ_SUPP_CONTACT_ADDRESSES_INT`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `SupplierContactImportTemplate / POZ_SUPP_CONTACT_ADDRESSES_INT` — name=NONE, keys=100%, cols=57% → L
  - `PONNegotiationLinesImportTemplate / ExternalCostFactors` — name=NONE, keys=100%, cols=29% → L
  - `PONNegotiationLinesImportTemplate / QuantityBasedPriceTiers` — name=NONE, keys=100%, cols=29% → L
  - `PONNegotiationLinesImportTemplate / PriceBreaks` — name=NONE, keys=100%, cols=29% → L
  - `SupplierBusinessClassificationImportTemplate / POZ_SUP_BUS_CLASS_INT` — name=NONE, keys=100%, cols=29% → L

### T_PROJ_ENT_RES_INTERFACE (prefix: TH9) — NEEDS_REVIEW
- **Prior:** YES → `ProjectEnterpriseExpenseResourcesImportTemplate / PJT_PRJ_ENT_RES_INTERFACE; ProjectEnterpriseResourcesImportTemplate / PJT_PRJ_ENT_RES_INTERFACE`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `ProjectEnterpriseResourcesImportTemplate / PJT_PRJ_ENT_RES_INTERFACE` — name=NONE, keys=100%, cols=50% → L
  - `SupplierContactImportTemplate / POZ_SUPP_CONTACT_ADDRESSES_INT` — name=NONE, keys=100%, cols=15% → L
  - `CustomerImportTemplate / HZ_IMP_CONTACTPTS_T` — name=NONE, keys=100%, cols=5% → L
  - `CustomerImportTemplate / RA_CUSTOMER_BANKS_INT_ALL` — name=NONE, keys=100%, cols=5% → L
  - `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=100%, cols=5% → L

### T_PROJ_RES_REQ_INTERFACE (prefix: TI1) — NEEDS_REVIEW
- **Prior:** YES → `ProjectResourceRequestImportTemplate / PJR_RES_REQ_INTERFACE`
- **Decision:** Prior references file/tab not found in 26B catalog or below all thresholds
- **Candidates evaluated:**
  - `ProjectAssetProcessingImportTemplate / Project Assets` — name=NONE, keys=50%, cols=5% → L
  - `ProjectUnprocessedLaborExpenditureItemImportTemplate / PJC_TXN_XFACE_STAGE_ALL` — name=NONE, keys=50%, cols=5% → L
  - `PayablesStandardInvoiceImportTemplate / AP_INVOICE_LINES_INTERFACE` — name=NONE, keys=50%, cols=4% → L
  - `ProjectAssetProcessingImportTemplate / Project Asset Assignments` — name=NONE, keys=50%, cols=4% → L
  - `ProjectUnprocessedExpenseReportExpenditureItemImportTemplate / PJC_TXN_XFACE_STAGE_ALL` — name=NONE, keys=50%, cols=4% → L

### T_SCP_SALESORDER (prefix: TK1) — NEEDS_REVIEW
- **Prior:** YES → `ScpSalesOrderImportTemplate / SalesOrder_`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ScpSalesOrderImportTemplate / SalesOrder_` — name=NONE, keys=100%, cols=97% → L
  - `ScpPurchaseOrderRequisitionImportTemplate / PurchaseOrderRequisition_` — name=NONE, keys=100%, cols=57% → L
  - `ScpTransferOrderImportTemplate / TransferOrder_` — name=NONE, keys=100%, cols=55% → L
  - `ScpWorkOrderSuppliesImportTemplate / WorkOrderSupplies_` — name=NONE, keys=100%, cols=53% → L
  - `ScpExternalForecastImportTemplate / ExternalForecast_` — name=NONE, keys=100%, cols=18% → L

### T_WIE_INT_BATCHES_VL (prefix: TD5) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkDefinitionTemplate / Import Batches; ProcessWorkDefinitionTemplate / Import Batches; WorkDefinitionTemplate / Import Batches`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `CseGenealogyBulkImport / Import Batches` — name=NONE, keys=100%, cols=100% → L
  - `CseInstalledBaseAssetImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `CseMeterReadingsImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `CseWarrantyCoverageImport / Batches` — name=NONE, keys=100%, cols=100% → L
  - `MaintenanceWorkDefinitionTemplate / Import Batches` — name=NONE, keys=100%, cols=100% → L

### T_WIS_WD_DETAILS_INT-ARES (prefix: TD11) — NEEDS_REVIEW
- **Prior:** YES → `ProcessWorkDefinitionTemplate / Operation Alternate Resources; WorkDefinitionTemplate / Operation Alternate Resources`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `ProcessWorkDefinitionTemplate / Operation Alternate Resources` — name=NONE, keys=100%, cols=100% → L
  - `WorkDefinitionTemplate / Operation Alternate Resources` — name=NONE, keys=100%, cols=100% → L
  - `ProcessWorkDefinitionTemplate / Operation Resources` — name=NONE, keys=100%, cols=83% → L
  - `WorkDefinitionTemplate / Operation Resources` — name=NONE, keys=100%, cols=83% → L
  - `MaintenanceWorkDefinitionTemplate / Operations Resources` — name=NONE, keys=100%, cols=75% → L

### T_WIS_WD_DETAILS_INT-ATO-CMP (prefix: TD9) — NEEDS_REVIEW
- **Prior:** YES → `WorkDefinitionTemplate / Operation Items - ATO Model`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `WorkDefinitionTemplate / Operation Items - ATO Model` — name=NONE, keys=100%, cols=100% → L
  - `ProcessWorkDefinitionTemplate / Operation Items` — name=NONE, keys=100%, cols=99% → L
  - `WorkDefinitionTemplate / Operation Items - Standard` — name=NONE, keys=100%, cols=99% → L
  - `MaintenanceWorkDefinitionTemplate / Operations Materials` — name=NONE, keys=100%, cols=97% → L
  - `ProcessWorkDefinitionTemplate / Operation Outputs` — name=NONE, keys=100%, cols=65% → L

### T_WIS_WD_DETAILS_INT-STNDCMP (prefix: TD8) — NEEDS_REVIEW
- **Prior:** YES → `WorkDefinitionTemplate / Operation Items - Standard`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `ProcessWorkDefinitionTemplate / Operation Items` — name=NONE, keys=100%, cols=73% → L
  - `WorkDefinitionTemplate / Operation Items - Standard` — name=NONE, keys=100%, cols=73% → L
  - `MaintenanceWorkDefinitionTemplate / Operations Materials` — name=NONE, keys=100%, cols=68% → L
  - `WorkDefinitionTemplate / Operation Items - ATO Model` — name=NONE, keys=100%, cols=65% → L
  - `MaintenanceWorkOrderTemplate / Operation Materials` — name=NONE, keys=100%, cols=38% → L

### T_WIS_WORK_DEFINITIONS_INT (prefix: TD6) — NEEDS_REVIEW
- **Prior:** YES → `ProcessWorkDefinitionTemplate / Work Definition Headers`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Header` — name=NONE, keys=100%, cols=64% → M
  - `MaintenanceWorkOrderTemplate / Work orders` — name=NONE, keys=100%, cols=62% → M
  - `ProcessWorkOrderTemplate / Work Order Header` — name=NONE, keys=100%, cols=14% → M
  - `WorkOrderMaterialTransactionTemplate / Material Transaction Header` — name=NONE, keys=100%, cols=10% → M
  - `ProcessWorkOrderMaterialTransactionTemplate / Material Transaction Header` — name=NONE, keys=100%, cols=7% → M

### T_WO_ASSEMBLY_COMPONENT (prefix: TM1) — NEEDS_REVIEW
- **Prior:** YES → `WorkOrderTemplate / Work Order Assembly Component`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `MaintenanceWorkOrderTemplate / Operation Materials` — name=NONE, keys=100%, cols=98% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Materials` — name=NONE, keys=100%, cols=98% → L
  - `WorkOrderTemplate / Work Order Operation Materials` — name=NONE, keys=100%, cols=98% → L
  - `WorkOrderTemplate / Work Order Assembly Component` — name=NONE, keys=100%, cols=98% → L
  - `MaintenanceWorkDefinitionTemplate / Work Definitions` — name=NONE, keys=100%, cols=67% → L

### T_WO_BATCHES (prefix: TI5) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Import Batches; ProcessWorkOrderTemplate / Work Order Batches; WorkOrderTemplate / Work Order Batches`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `CseInstalledBaseAssetImport / Batches` — name=NONE, keys=100%, cols=67% → L
  - `CseMeterReadingsImport / Batches` — name=NONE, keys=100%, cols=67% → L
  - `CseWarrantyCoverageImport / Batches` — name=NONE, keys=100%, cols=67% → L
  - `MaintenanceWorkOrderTemplate / Import Batches` — name=NONE, keys=100%, cols=67% → L
  - `MntMaintenanceProgramImport / Batches` — name=NONE, keys=100%, cols=67% → L

### T_WO_HEADER (prefix: TI6) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Work orders; ProcessWorkOrderTemplate / Work Order Header; WorkOrderTemplate / Work Order Header`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Header` — name=NONE, keys=100%, cols=96% → L
  - `MaintenanceWorkOrderTemplate / Work orders` — name=NONE, keys=100%, cols=89% → L
  - `ProcessWorkOrderTemplate / Work Order Header` — name=NONE, keys=100%, cols=74% → L
  - `WorkOrderOperationTransactionTemplate / Operation Transaction Header` — name=NONE, keys=100%, cols=57% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Outputs` — name=NONE, keys=100%, cols=55% → L

### T_WO_MATERIAL_LOT_NUMBERS (prefix: TM3) — NEEDS_REVIEW
- **Prior:** YES → `WorkOrderTemplate / Work Order Material Lot Numbers`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `CseInstalledBaseAssetImport / Assets` — name=NONE, keys=100%, cols=80% → L
  - `ProcessWorkOrderTemplate / Work Order Product Lots` — name=NONE, keys=100%, cols=80% → L
  - `WorkOrderTemplate / Work Order Product Lot Numbers` — name=NONE, keys=100%, cols=80% → L
  - `WorkOrderTemplate / Work Order Serial Numbers` — name=NONE, keys=100%, cols=80% → L
  - `WorkOrderTemplate / Work Order Matl Serial Numbers` — name=NONE, keys=100%, cols=80% → L

### T_WO_MATL_SERIAL_NUMBER (prefix: TM2) — NEEDS_REVIEW
- **Prior:** YES → `WorkOrderTemplate / Work Order Matl Serial Numbers`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Serial Numbers` — name=NONE, keys=100%, cols=100% → L
  - `WorkOrderTemplate / Work Order Matl Serial Numbers` — name=NONE, keys=100%, cols=100% → L
  - `CseInstalledBaseAssetImport / Assets` — name=NONE, keys=100%, cols=75% → L
  - `CseWarrantyCoverageImport / Warranty Coverages` — name=NONE, keys=100%, cols=75% → L
  - `CseWarrantyCoverageImport / Covered Items` — name=NONE, keys=100%, cols=75% → L

### T_WO_OPERATIONS (prefix: TI7) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Work Order Operations; ProcessWorkOrderTemplate / Work Order Operation Outputs; ProcessWorkOrderTemplate / Work Order Operations; WorkOrderTemplate / Work Order Operation Outputs; WorkOrderTemplate / Work Order Operations`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Operations` — name=NONE, keys=100%, cols=97% → L
  - `ProcessWorkOrderTemplate / Work Order Operations` — name=NONE, keys=100%, cols=95% → L
  - `ProcessWorkDefinitionTemplate / Work Definition Operations` — name=NONE, keys=100%, cols=89% → L
  - `WorkDefinitionTemplate / Work Definition Operations` — name=NONE, keys=100%, cols=89% → L
  - `MaintenanceWorkDefinitionTemplate / Work Definition Operations` — name=NONE, keys=100%, cols=85% → L

### T_WO_OPERATION_MATERIALS (prefix: TJ1) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Operation Materials; ProcessWorkOrderTemplate / Work Order Operation Materials; WorkOrderTemplate / Work Order Operation Materials`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Operation Materials` — name=NONE, keys=100%, cols=99% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Materials` — name=NONE, keys=100%, cols=97% → L
  - `MaintenanceWorkOrderTemplate / Operation Materials` — name=NONE, keys=100%, cols=96% → L
  - `WorkOrderTemplate / Work Order Assembly Component` — name=NONE, keys=100%, cols=84% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Outputs` — name=NONE, keys=100%, cols=62% → L

### T_WO_OPERATION_RESOURCES (prefix: TI9) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Operation resources; ProcessWorkOrderTemplate / Work Order Operation Resources; WorkOrderTemplate / Work Order Operation Resources`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `MaintenanceWorkOrderTemplate / Operation resources` — name=NONE, keys=100%, cols=97% → L
  - `WorkOrderTemplate / Work Order Operation Resources` — name=NONE, keys=100%, cols=97% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Resources` — name=NONE, keys=100%, cols=95% → L
  - `MaintenanceWorkDefinitionTemplate / Operations Resources` — name=NONE, keys=100%, cols=82% → L
  - `ProcessWorkDefinitionTemplate / Operation Resources` — name=NONE, keys=100%, cols=82% → L

### T_WO_OPN_RES_INSTANCES (prefix: TJ0) — NEEDS_REVIEW
- **Prior:** YES → `MaintenanceWorkOrderTemplate / Operation Resource instances; ProcessWorkOrderTemplate / Work Order Opn Res Instances; WorkOrderTemplate / Work Order Opn Res Instances`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `MaintenanceWorkOrderTemplate / Operation Resource instances` — name=NONE, keys=100%, cols=98% → L
  - `ProcessWorkOrderTemplate / Work Order Opn Res Instances` — name=NONE, keys=100%, cols=98% → L
  - `WorkOrderTemplate / Work Order Opn Res Instances` — name=NONE, keys=100%, cols=98% → L
  - `MaintenanceWorkOrderTemplate / Operation resources` — name=NONE, keys=100%, cols=94% → L
  - `ProcessWorkOrderTemplate / Work Order Operation Resources` — name=NONE, keys=100%, cols=94% → L

### T_WO_PRODUCT_LOT_NUMBERS (prefix: TJ2) — NEEDS_REVIEW
- **Prior:** YES → `ProcessWorkOrderTemplate / Work Order Product Lots; WorkOrderTemplate / Work Order Product Lot Numbers`
- **Decision:** Multi-mapping contested — see audit.md for per-leg evidence
- **Candidates evaluated:**
  - `ProcessWorkOrderTemplate / Work Order Product Lots` — name=NONE, keys=100%, cols=100% → L
  - `WorkOrderTemplate / Work Order Product Lot Numbers` — name=NONE, keys=100%, cols=100% → L
  - `WorkOrderTemplate / Work Order Material Lot Numbers` — name=NONE, keys=100%, cols=67% → L
  - `CseInstalledBaseAssetImport / Assets` — name=NONE, keys=100%, cols=50% → L
  - `CseWarrantyCoverageImport / Warranty Coverages` — name=NONE, keys=100%, cols=50% → L

### T_WO_SERIAL_NUMBERS (prefix: TJ3) — NEEDS_REVIEW
- **Prior:** YES → `WorkOrderTemplate / Work Order Serial Numbers`
- **Decision:** Prior claim scores Low against 26B catalog — verify
- **Candidates evaluated:**
  - `WorkOrderTemplate / Work Order Serial Numbers` — name=NONE, keys=100%, cols=100% → L
  - `CseInstalledBaseAssetImport / Assets` — name=NONE, keys=100%, cols=98% → L
  - `MaintenanceWorkDefinitionTemplate / Work Definitions` — name=NONE, keys=100%, cols=98% → L
  - `MaintenanceWorkDefinitionTemplate / Work Definition Operations` — name=NONE, keys=100%, cols=98% → L
  - `MaintenanceWorkDefinitionTemplate / Operations Materials` — name=NONE, keys=100%, cols=98% → L

### T_BANKS_BRANCHES (prefix: T32) — NEEDS_REVIEW
- **Prior:** YES → `RapidImplementationForCashManagement / Bank Account`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `RapidImplementationForCashManagement / Bank Account` — name=NONE, keys=100%, cols=96% → M
  - `PayablesPaymentRequestImportTemplate / AP_PAYMENT_REQUESTS_INT` — name=NONE, keys=50%, cols=30% → L
  - `CashManagementBankStatementImportTemplate / Statement Headers` — name=NONE, keys=50%, cols=9% → L

### T_BPA_PO_HEADERS_INTERFACE (prefix: T62) — NEEDS_REVIEW
- **Prior:** YES → `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=50%, cols=62% → M
  - `POContractPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=50%, cols=59% → M
  - `SchExternalPurchasePricesImportTemplate / SCH_EPP_HEADERS_INT` — name=NONE, keys=50%, cols=2% → L
  - `POBlanketPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE` — name=NONE, keys=50%, cols=1% → L
  - `POContractPurchaseAgreementImportTemplate / PO_GA_ORG_ASSIGN_INTERFACE` — name=NONE, keys=50%, cols=1% → L

### T_INV_SERIAL_NUMBERS_INT (prefix: T49) — NEEDS_REVIEW
- **Prior:** UNMAPPED → `InventoryTransactionImportTemplate / INV_SERIAL_NUMBERS_INTERFACE`
- **Decision:** Potential new mapping — Medium confidence; verify with Brad
- **Candidates evaluated:**
  - `InventoryTransactionImportTemplate / INV_SERIAL_NUMBERS_INTERFACE` — name=PARTIAL, keys=0%, cols=98% → M
  - `ReceivingReceiptImportTemplate / INV_SERIAL_NUMBERS_INTERFACE` — name=PARTIAL, keys=0%, cols=72% → M
  - `InterfacedPickTransactionsImportTemplate / INV_SERIAL_NUMBERS_INTERFACE` — name=PARTIAL, keys=0%, cols=1% → M
  - `PerformShippingTransactionImportTemplate / INV_SERIAL_NUMBERS_INTERFACE` — name=PARTIAL, keys=0%, cols=0% → M
  - `InventoryTransactionImportTemplate / INV_TRANSACTION_LOTS_INTERFACE` — name=NONE, keys=0%, cols=83% → L

## Changed From Prior

### T_AWARD_KEYWORDS (prefix: TE5) — YES
- **Prior:** YES → `ImportAwards / Award Keywords`
- **Decision:** Collapsed from multi — 1/2 legs scored H; rest below threshold
- **Candidates evaluated:**
  - `ImportAwards / Award Keywords` — name=EXACT, keys=100%, cols=100% → H
  - `ImportAwards / Award Projects` — name=NONE, keys=100%, cols=67% → L
  - `ImportAwards / Award Project Funding Sources` — name=NONE, keys=100%, cols=67% → L
  - `ImportAwards / Award References` — name=NONE, keys=100%, cols=67% → L
  - `ImportAwards / Award Certifications` — name=NONE, keys=100%, cols=67% → L
- **Note:** Prefix mismatch — expected T_<tab> convention, got prefix=TE5 for tab=Award Keywords

### T_IBY_TEMP_EXT_BANK_ACCT (prefix: T15) — YES
- **Prior:** YES → `SupplierBankAccountImportTemplate / IBY_TEMP_EXT_BANK_ACCTS`
- **Decision:** Collapsed from multi — 1/2 legs scored M; rest below threshold
- **Candidates evaluated:**
  - `SupplierBankAccountImportTemplate / IBY_TEMP_EXT_BANK_ACCTS` — name=PARTIAL, keys=0%, cols=0% → M
  - `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=12%, cols=51% → M
  - `POContractPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=12%, cols=51% → M
  - `InventoryTransactionImportTemplate / INV_TRANSACTIONS_INTERFACE` — name=NONE, keys=12%, cols=44% → M
  - `ReceivingReceiptImportTemplate / RCV_TRANSACTIONS_INTERFACE` — name=NONE, keys=12%, cols=44% → M
- **Note:** Prefix mismatch — expected T_<tab> convention, got prefix=T15 for tab=IBY_TEMP_EXT_BANK_ACCTS

### T_BPA_PO_LINES_INTERFACE (prefix: T64) — UNMAPPED
- **Prior:** YES → `(none)`
- **Decision:** No FBDI tab in 26B catalog scores above threshold
- **Candidates evaluated:**
  - `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=33%, cols=36% → L
  - `POContractPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=33%, cols=36% → L
  - `POBlanketPurchaseAgreementImportTemplate / PO_LINES_INTERFACE` — name=NONE, keys=0%, cols=50% → L
  - `POPurchaseOrderImportTemplate / PO_LINES_INTERFACE` — name=NONE, keys=0%, cols=50% → L
  - `RequisitionImportTemplate / POR_REQ_LINES_INTERFACE_ALL` — name=NONE, keys=0%, cols=35% → L

### T_BPA_PO_LINE_LOCATIONS (prefix: T63) — UNMAPPED
- **Prior:** YES → `(none)`
- **Decision:** No FBDI tab in 26B catalog scores above threshold
- **Candidates evaluated:**
  - `POBlanketPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=0%, cols=87% → L
  - `POContractPurchaseAgreementImportTemplate / PO_HEADERS_INTERFACE` — name=NONE, keys=0%, cols=87% → L
  - `POBlanketPurchaseAgreementImportTemplate / PO_LINE_LOCATIONS_INTERFACE` — name=NONE, keys=0%, cols=85% → L
  - `PONNegotiationLinesImportTemplate / LinesOrderOutcome` — name=NONE, keys=0%, cols=85% → L
  - `POPurchaseOrderImportTemplate / PO_LINES_INTERFACE` — name=NONE, keys=0%, cols=85% → L

### T_DOO_ORDER_LINES_EFF_B (prefix: TC6) — UNMAPPED
- **Prior:** YES → `(none)`
- **Decision:** No FBDI tab in 26B catalog scores above threshold
- **Candidates evaluated:**
  - `ChangeOrderImportTemplate / EGO_CHANGE_EFF_INT` — name=NONE, keys=0%, cols=91% → L
  - `ChangeOrderImportTemplate / EGO_ITEM_INTF_EFF_B` — name=NONE, keys=0%, cols=91% → L
  - `ItemImportTemplate / EGO_ITEM_INTF_EFF_B` — name=NONE, keys=0%, cols=91% → L
  - `ItemImportTemplate / EGO_ITEM_REVISION_INTF_EFF_B` — name=NONE, keys=0%, cols=91% → L
  - `ItemImportTemplate / EGO_ITEM_SUPPLIER_INTF_EFF_B` — name=NONE, keys=0%, cols=91% → L

## Prefix Mismatches

| Applaud Table | Prefix | Notes |
|---|---|---|
| T_ADDITIONALTRANSFERORDERCOST | TL2 | Prefix mismatch — expected T_<tab> convention, got prefix=TL2 for tab=Additional Transfer Order Costs |
| T_AWARD_BUDGET_PERIODS | TF2 | Prefix mismatch — expected T_<tab> convention, got prefix=TF2 for tab=Award Budget Periods |
| T_AWARD_CERTIFICATIONS | TE7 | Prefix mismatch — expected T_<tab> convention, got prefix=TE7 for tab=Award Certifications |
| T_AWARD_FUNDING | TF5 | Prefix mismatch — expected T_<tab> convention, got prefix=TF5 for tab=Award Funding |
| T_AWARD_FUNDING_ALLOCATIONS | TF6 | Prefix mismatch — expected T_<tab> convention, got prefix=TF6 for tab=Award Funding Allocations |
| T_AWARD_FUNDING_SOURCE | TE2 | Prefix mismatch — expected T_<tab> convention, got prefix=TE2 for tab=Award Funding Sources |
| T_AWARD_KEYWORDS | TE5 | Prefix mismatch — expected T_<tab> convention, got prefix=TE5 for tab=Award Keywords |
| T_AWARD_ORGANIZATION_CREDITS | TF3 | Prefix mismatch — expected T_<tab> convention, got prefix=TF3 for tab=Award Organization Credits |
| T_AWARD_PERSONNEL | TF4 | Prefix mismatch — expected T_<tab> convention, got prefix=TF4 for tab=Award Personnel |
| T_AWARD_PROJECTS | TE3 | Prefix mismatch — expected T_<tab> convention, got prefix=TE3 for tab=Award Projects |
| T_AWARD_PROJECT_FUNDING_SOURCE | TE4 | Prefix mismatch — expected T_<tab> convention, got prefix=TE4 for tab=Award Project Funding Sources |
| T_AWARD_REFERENCES | TE6 | Prefix mismatch — expected T_<tab> convention, got prefix=TE6 for tab=Award References |
| T_AWARD_TERMS_AND_CONDITIONS | TE8 | Prefix mismatch — expected T_<tab> convention, got prefix=TE8 for tab=Award Terms and Conditions |
| T_CST_INTR_STD_COST_DETAIL | T83 | Prefix mismatch — expected T_<tab> convention, got prefix=T83 for tab=CST_INTERFACE_STD_COST_DETAILS |
| T_CST_INTR_STD_COST_HEADERS | T84 | Prefix mismatch — expected T_<tab> convention, got prefix=T84 for tab=CST_INTERFACE_STD_COST_HEADERS |
| T_DOO_ORDER_CHARGE_COMPS | TC1 | Prefix mismatch — expected T_<tab> convention, got prefix=TC1 for tab=DOO_ORDER_CHARGE_COMPS_INT |
| T_DOO_ORDER_HDRS_ALL_EFF_B | TC3 | Prefix mismatch — expected T_<tab> convention, got prefix=TC3 for tab=DOO_ORDER_HDRS_ALL_EFF_B_INT |
| T_DOO_ORDER_HEADERS_ALL | TC4 | Prefix mismatch — expected T_<tab> convention, got prefix=TC4 for tab=DOO_ORDER_HEADERS_ALL_INT |
| T_DOO_ORDER_LINES_ALL | TC5 | Prefix mismatch — expected T_<tab> convention, got prefix=TC5 for tab=DOO_ORDER_LINES_ALL_INT |
| T_EGP_ITEM_CATEGORIES_INT | T87 | Prefix mismatch — expected T_<tab> convention, got prefix=T87 for tab=EGP_ITEM_CATEGORIES_INTERFACE |
| T_EGP_ITEM_REVISION_INT | T89 | Prefix mismatch — expected T_<tab> convention, got prefix=T89 for tab=EGP_ITEM_REVISIONS_INTERFACE |
| T_EGP_TRADING_PARTNER_ITEMS | T85 | Prefix mismatch — expected T_<tab> convention, got prefix=T85 for tab=EGP_TRADING_PARTNER_ITEMS_INTF |
| T_FA_ADJUSTMENTS | T68 | Prefix mismatch — expected T_<tab> convention, got prefix=T68 for tab=FA_ADJUSTMENTS_T |
| T_GL_BUDGETS_INTERFACE | TF1 | Prefix mismatch — expected T_<tab> convention, got prefix=TF1 for tab=GL_BUDGET_INTERFACE |
| T_GL_SEGMENT_VALUES_INT | T02 | Prefix mismatch — expected T_<tab> convention, got prefix=T02 for tab=GL_SEGMENT_VALUES_INTERFACE |
| T_HZ_IMP_ACCOUNTRELS_T | T27 | Prefix mismatch — expected T_<tab> convention, got prefix=T27 for tab=HZ_IMP_ACCOUNTRELS |
| T_HZ_IMP_PARTY_SITES_T | T52 | Prefix mismatch — expected T_<tab> convention, got prefix=T52 for tab=HZ_IMP_PARTYSITES_T |
| T_IBY_TEMP_EXT_BANK_ACCT | T15 | Prefix mismatch — expected T_<tab> convention, got prefix=T15 for tab=IBY_TEMP_EXT_BANK_ACCTS |
| T_INV_TRANSACTION_LOTS_INT | T48 | Prefix mismatch — expected T_<tab> convention, got prefix=T48 for tab=Material Transaction Lots |
| T_MSC_ST_ASSIGNMENT_SETS | T04 | Prefix mismatch — expected T_<tab> convention, got prefix=T04 for tab=AssignmentSets_ |
| T_MSC_ST_MEASURE_DATA | T05 | Prefix mismatch — expected T_<tab> convention, got prefix=T05 for tab=DPForecasts_ |
| T_MSC_ST_MEASURE_DATA_BOOKINGS | T06 | Prefix mismatch — expected T_<tab> convention, got prefix=T06 for tab=BookingHistory_ |
| T_MSC_ST_SOURCING_RULES | T03 | Prefix mismatch — expected T_<tab> convention, got prefix=T03 for tab=SourcingRules_ |
| T_POZ_SUP_CONTACTS_INT | T11 | Prefix mismatch — expected T_<tab> convention, got prefix=T11 for tab=POZ_SUP_CONTACTS |
| T_PROJECT_CLASSIFICATIONS | T56 | Prefix mismatch — expected T_<tab> convention, got prefix=T56 for tab=Project Classifications |
| T_PROJECT_TEAM_MEMBERS | T57 | Prefix mismatch — expected T_<tab> convention, got prefix=T57 for tab=Project Team Members |
| T_RA_INTERFACE_DISTRIBUTIONS | TA3 | Prefix mismatch — expected T_<tab> convention, got prefix=TA3 for tab=RA_INTERFACE_DISTRIBUTIONS_ALL |
| T_REFERENCE_ACCOUNTS | T29 | Prefix mismatch — expected T_<tab> convention, got prefix=T29 for tab=Reference Accounts |
| T_SAFETYSTOCKLEVEL | TK9 | Prefix mismatch — expected T_<tab> convention, got prefix=TK9 for tab=SafetyStockLevel_ |
| T_SCP_EXTERNALFORECAST | TK2 | Prefix mismatch — expected T_<tab> convention, got prefix=TK2 for tab=ExternalForecast_ |
| T_SCP_UOM_CONVERSION | TM4 | Prefix mismatch — expected T_<tab> convention, got prefix=TM4 for tab=UOMConversion_ |
| T_TRANSFERORDERLINES | TL1 | Prefix mismatch — expected T_<tab> convention, got prefix=TL1 for tab=Transfer Order Lines |
| T_WIS_WD_DETAILS_INT-OPS | TD7 | Prefix mismatch — expected T_<tab> convention, got prefix=TD7 for tab=Work Definition Operations |
| T_WIS_WD_DETAILS_INT-RSC | TD10 | Prefix mismatch — expected T_<tab> convention, got prefix=TD10 for tab=Operation Resources |
