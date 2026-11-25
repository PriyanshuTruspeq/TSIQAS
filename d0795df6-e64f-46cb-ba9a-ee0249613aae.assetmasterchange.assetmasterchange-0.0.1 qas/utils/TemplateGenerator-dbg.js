sap.ui.define([
    "assetmasterchange/utils/XlsxFormatter"
], function (XlsxFormatter) {
    "use strict";

    /**
     * AssetMasterChangeTemplateGenerator v3.0
     *
     * BACKWARD COMPATIBLE with v2.6 – all existing scenarios, sheets, and field patterns preserved.
     *
     * NEW IN v3.0:
     *   • Dynamic LEDGER_AREAS – pass options.ledgerAreas to support any chart of depreciation
     *   • LEDGER_ACTION_CODE_LEDGERS – only 0L gets Ledger_ActionCode (2L/3L do NOT per EDMX)
     *   • New per-area fields from EDMX: NegativeAmountIsAllowed, TimeBasedValuation_ActionCode,
     *     AssetDeactivationDate, SpecialDeprStartDate, InterestCalcStartDate,
     *     VariableDeprPercent, DeprCalcBaseValuePercent, ShiftOperationFactor,
     *     IsShutDown, AssetRevaluationIndex
     *   • New account assignment fields: Segment, AssetLocation, VehicleLicensePlateNumber,
     *     TaxJurisdiction, BusinessPlace, Fund, GrantID, FunctionalArea
     *   • Scenario 9: Full payload example mirroring exact SAP OData structure
     *   • New "Ledger_Areas" reference sheet
     *
     * LEDGER_AREA FORMAT: "<Ledger>_<DepreciationArea>"  e.g. "0L_01", "0L_15", "2L_32", "3L_34"
     *
     * Only ledgers in LEDGER_ACTION_CODE_LEDGERS receive a Ledger_ActionCode column.
     * Ledger 0L carries AssetCapitalizationDate at ledger level → needs action code.
     * Ledger 2L / 3L have no ledger-level FixedAssetObjectActionCode in the SAP payload.
     */
    return class AssetMasterChangeTemplateGenerator {

        constructor(options = {}) {
            this._xlsxFormatter = options.xlsxFormatter || new XlsxFormatter();
            this._customProcessor = options.customExcelProcessor;

            /**
             * ALL supported ledger+area combinations.
             * Override via options.ledgerAreas to match your system's chart of depreciation.
             *
             * Default: Standard S/4HANA India configuration
             *   0L_01  Book depreciation        (Ledger 0L, Area 01)
             *   0L_15  Income-tax depreciation   (Ledger 0L, Area 15)
             *   2L_32  Consolidation (IFRS)      (Ledger 2L, Area 32)
             *   3L_34  Special purpose           (Ledger 3L, Area 34)
             *
             * Example to extend:
             *   new AssetMasterChangeTemplateGenerator({
             *     ledgerAreas: ["0L_01","0L_10","0L_15","2L_32","3L_34"]
             *   })
             */
            this.LEDGER_AREAS = options.ledgerAreas || [
                "0L_01",
                "0L_15",
                "2L_32",
                "3L_34"
            ];

            /**
             * Only these ledger codes get a Ledger_ActionCode column.
             * Ledger 0L carries AssetCapitalizationDate at ledger level.
             * Ledgers 2L and 3L have no FixedAssetObjectActionCode at ledger level.
             */
            this.LEDGER_ACTION_CODE_LEDGERS = options.ledgerActionCodeLedgers || ["0L"];

            // Pre-compute unique ledger codes
            this._ledgerCodes = [...new Set(this.LEDGER_AREAS.map(la => la.split("_")[0]))];
        }

        // ─────────────────────────────────────────────────────────────
        // PUBLIC API
        // ─────────────────────────────────────────────────────────────

        /**
         * Generates Asset Master Change template
         * @returns {Promise<Blob>} Promise resolving with the change template Blob
         */
        generateChangeTemplateBlob() {
            return new Promise(async (resolve, reject) => {
                try {
                    let blob = null;

                    if (this._customProcessor && typeof this._customProcessor.createChangeTemplateBlob === 'function') {
                        try {
                            console.debug("TemplateGenerator: Using custom processor for change template.");
                            blob = await this._customProcessor.createChangeTemplateBlob();
                            if (!(blob instanceof Blob)) blob = null;
                        } catch (customError) {
                            console.warn("TemplateGenerator: Custom processor failed, using default.", customError);
                            blob = null;
                        }
                    }

                    if (!blob) {
                        console.debug("TemplateGenerator: Generating comprehensive Asset Master Change template.");
                        const wb = await this._createEnhancedWorkbook();
                        blob = await this._xlsxFormatter.workbookToBlob(wb);
                    }

                    resolve(blob);
                } catch (error) {
                    console.error("Error generating change template blob:", error);
                    reject(error);
                }
            });
        }

        /**
         * Gets change template filename
         * @returns {string} The template filename
         */
        getChangeTemplateFileName() {
            let defaultName = "Asset_Master_Change_Template_v3.0_Enhanced.xlsx";
            if (this._customProcessor && typeof this._customProcessor.getChangeTemplateFileName === 'function') {
                try {
                    return this._customProcessor.getChangeTemplateFileName() || defaultName;
                } catch (e) {}
            }
            return defaultName;
        }

        // ─────────────────────────────────────────────────────────────
        // WORKBOOK CONSTRUCTION
        // ─────────────────────────────────────────────────────────────

        /**
         * Creates enhanced workbook with instructions and data sheets
         * @private
         */
        async _createEnhancedWorkbook() {
            const wb = window.XLSX.utils.book_new();

            // Instructions sheet
            const instructionsWS = window.XLSX.utils.aoa_to_sheet(this._getInstructionsData());
            this._styleInstructionsSheet(instructionsWS);
            window.XLSX.utils.book_append_sheet(wb, instructionsWS, "Instructions");

            // Change_Data sheet with comprehensive example rows
            const headers = this._getAssetMasterChangeHeaders();
            const exampleRows = this._getAllExampleRows();
            const changeDataWS = await this._createDataSheet(headers, exampleRows);
            window.XLSX.utils.book_append_sheet(wb, changeDataWS, "Change_Data");

            // Action_Codes sheet
            const actionCodesWS = window.XLSX.utils.aoa_to_sheet(this._getActionCodesReferenceData());
            this._styleReferenceSheet(actionCodesWS);
            window.XLSX.utils.book_append_sheet(wb, actionCodesWS, "Action_Codes");

            // Field_Reference sheet
            const fieldReferenceWS = window.XLSX.utils.aoa_to_sheet(this._getFieldReferenceData());
            this._styleReferenceSheet(fieldReferenceWS);
            window.XLSX.utils.book_append_sheet(wb, fieldReferenceWS, "Field_Reference");

            // Scenarios sheet
            const scenariosWS = window.XLSX.utils.aoa_to_sheet(this._getScenariosExplanationData());
            this._styleReferenceSheet(scenariosWS);
            window.XLSX.utils.book_append_sheet(wb, scenariosWS, "Scenarios");

            // NEW: Ledger_Areas reference sheet
            const ledgerAreasWS = window.XLSX.utils.aoa_to_sheet(this._getLedgerAreasReferenceData());
            this._styleReferenceSheet(ledgerAreasWS);
            window.XLSX.utils.book_append_sheet(wb, ledgerAreasWS, "Ledger_Areas");

            return wb;
        }

        // ─────────────────────────────────────────────────────────────
        // DYNAMIC HEADER GENERATION
        // ─────────────────────────────────────────────────────────────

        /**
         * Returns complete ordered header array.
         * All original v2.6 fields preserved + new EDMX fields added.
         * Ledger/area columns are generated dynamically from this.LEDGER_AREAS.
         * @private
         * @returns {Array<string>}
         */
        _getAssetMasterChangeHeaders() {
            const headers = [
                // ── Required identification fields ─────────────────────
                "SequenceNumber", "CompanyCode", "FixedAsset",

                // ── General fields ─────────────────────────────────────
                "FixedAssetDescription_Value",      "FixedAssetDescription_Change",
                "AssetAdditionalDescription_Value", "AssetAdditionalDescription_Change",
                "AssetSerialNumber_Value",           "AssetSerialNumber_Change",

                // ── Account assignment fields (original v2.6) ──────────
                "WBSElementExternalID_Value",    "WBSElementExternalID_Change",
                "YY1_WBSELEMENT1_FAA_Value",     "YY1_WBSELEMENT1_FAA_Change",
                "Room_Value",                    "Room_Change",

                // ── Inventory fields ───────────────────────────────────
                "InventoryNote_Value", "InventoryNote_Change",

                // ── India-specific fields ──────────────────────────────
                "IN_AssetBlock_Value",            "IN_AssetBlock_Change",
                "IN_AssetPutToUseDate_Value",     "IN_AssetPutToUseDate_Change",
                "IN_AdditionalAssetBlock_Value",  "IN_AdditionalAssetBlock_Change",
                "IN_AssetIsResearchAndDev_Value", "IN_AssetIsResearchAndDev_Change",
                "IN_AssetIsPriorYear_Value",      "IN_AssetIsPriorYear_Change",

                // ── Account Assignment action + additional fields ───────
                "AccountAssignment_ActionCode",
                "AccountAssignment_ValidityStartDate",
                "CostCenter_Value",    "CostCenter_Change",
                "Plant_Value",         "Plant_Change",
                "ProfitCenter_Value",  "ProfitCenter_Change",
                // NEW v3.0: full FixedAssetAssignment_Type coverage
                "Segment_Value",                   "Segment_Change",
                "AssetLocation_Value",             "AssetLocation_Change",
                "VehicleLicensePlateNumber_Value", "VehicleLicensePlateNumber_Change",
                "TaxJurisdiction_Value",           "TaxJurisdiction_Change",
                "BusinessPlace_Value",             "BusinessPlace_Change",
                "Fund_Value",                      "Fund_Change",
                "GrantID_Value",                   "GrantID_Change",
                "FunctionalArea_Value",            "FunctionalArea_Change",

                // ── Additional Inventory fields ────────────────────────
                "Inventory_Value",          "Inventory_Change",
                "InventoryIsCounted_Value", "InventoryIsCounted_Change",

                // ── General depreciation date fallbacks ────────────────
                "AssetCapitalizationDate_Value",
                "DepreciationStartDate_Value",
                "ValidityStartDate_Value",
            ];

            // ── Per-area columns (dynamic for all configured areas) ────
            for (const la of this.LEDGER_AREAS) {
                headers.push(...this._getHeadersForArea(la));
            }

            return headers;
        }

        /**
         * Returns the ordered column list for a single ledger_area.
         * NEW v3.0: includes all EDMX fields not in v2.6.
         * @param {string} ledgerArea  e.g. "0L_01"
         * @private
         */
        _getHeadersForArea(ledgerArea) {
            const ledger = ledgerArea.split("_")[0];
            const cols = [];

            // ── Ledger-level (FixedAssetLedger_Type) ──────────────────
            // Only 0L gets Ledger_ActionCode (has FixedAssetObjectActionCode at ledger level)
            if (this.LEDGER_ACTION_CODE_LEDGERS.includes(ledger)) {
                cols.push(`Ledger_ActionCode_${ledgerArea}`);
            }
            cols.push(
                `AssetCapitalizationDate_${ledgerArea}_Value`,
                `AssetCapitalizationDate_${ledgerArea}_Change`,
                // NEW v3.0
                `AssetDeactivationDate_${ledgerArea}_Value`,
                `AssetDeactivationDate_${ledgerArea}_Change`
            );

            // ── Valuation-level (FixedAssetValuation_Type) ────────────
            cols.push(
                `Valuation_ActionCode_${ledgerArea}`,
                // NEW v3.0: NegativeAmountIsAllowed (boolean, no _Change pair)
                `NegativeAmountIsAllowed_${ledgerArea}`,
                `DepreciationStartDate_${ledgerArea}_Value`,
                `DepreciationStartDate_${ledgerArea}_Change`,
                // NEW v3.0
                `SpecialDeprStartDate_${ledgerArea}_Value`,
                `SpecialDeprStartDate_${ledgerArea}_Change`,
                `InterestCalcStartDate_${ledgerArea}_Value`,
                `InterestCalcStartDate_${ledgerArea}_Change`
            );

            // ── Time-based valuation (FixedAssetTimeBasedValuation_Type)
            cols.push(
                // NEW v3.0: independent TimeBasedValuation action code
                `TimeBasedValuation_ActionCode_${ledgerArea}`,
                `ValidityStartDate_${ledgerArea}_Value`,
                `DepreciationKey_${ledgerArea}_Value`,
                `DepreciationKey_${ledgerArea}_Change`,
                `PlannedUsefulLifeInYears_${ledgerArea}_Value`,
                `PlannedUsefulLifeInYears_${ledgerArea}_Change`,
                `PlannedUsefulLifeInPeriods_${ledgerArea}_Value`,
                `PlannedUsefulLifeInPeriods_${ledgerArea}_Change`,
                `ScrapAmountInCoCodeCrcy_${ledgerArea}_Value`,
                `ScrapAmountInCoCodeCrcy_${ledgerArea}_Change`,
                `AcqnProdnCostScrapPercent_${ledgerArea}_Value`,
                `AcqnProdnCostScrapPercent_${ledgerArea}_Change`,
                // NEW v3.0
                `VariableDeprPercent_${ledgerArea}_Value`,
                `VariableDeprPercent_${ledgerArea}_Change`,
                `DeprCalcBaseValuePercent_${ledgerArea}_Value`,
                `DeprCalcBaseValuePercent_${ledgerArea}_Change`,
                `ShiftOperationFactor_${ledgerArea}_Value`,
                `ShiftOperationFactor_${ledgerArea}_Change`,
                // NEW v3.0: boolean, no _Change pair
                `IsShutDown_${ledgerArea}`,
                `AssetRevaluationIndex_${ledgerArea}_Value`,
                `AssetRevaluationIndex_${ledgerArea}_Change`
            );

            return cols;
        }

        // ─────────────────────────────────────────────────────────────
        // EMPTY LEDGER FIELDS HELPER
        // ─────────────────────────────────────────────────────────────

        /**
         * Generate empty ledger fields for a specific area.
         * v3.0: now covers all new EDMX fields as well as original v2.6 fields.
         * @private
         */
        _getEmptyLedgerFields(ledgerArea) {
            const fields = {};
            for (const col of this._getHeadersForArea(ledgerArea)) {
                fields[col] = "";
            }
            return fields;
        }

        /**
         * Generate empty fields for ALL configured ledger areas.
         * @private
         */
        _getAllEmptyLedgerFields() {
            let fields = {};
            for (const la of this.LEDGER_AREAS) {
                fields = { ...fields, ...this._getEmptyLedgerFields(la) };
            }
            return fields;
        }

        // ─────────────────────────────────────────────────────────────
        // SCENARIO DEFINITIONS (v2.6 scenarios preserved exactly)
        // ─────────────────────────────────────────────────────────────

        /**
         * ✅ SCENARIO 1: Simple CHANGE - Update depreciation key only
         * Use Case: Correcting wrong depreciation key for existing asset
         * @private
         */
        _getScenario1_SimpleChange() {
            return {
                "SequenceNumber": 1,
                "CompanyCode": "1000",
                "FixedAsset": "000000100001",

                "FixedAssetDescription_Value": "",
                "FixedAssetDescription_Change": "",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                // ✅ CHANGE Area 0L_01 only - Change depreciation key
                "Ledger_ActionCode_0L_01": "02",
                "Valuation_ActionCode_0L_01": "02",
                "TimeBasedValuation_ActionCode_0L_01": "02",
                "NegativeAmountIsAllowed_0L_01": "true",
                "AssetCapitalizationDate_0L_01_Value": "2026-04-01",
                "AssetCapitalizationDate_0L_01_Change": "Y",
                "DepreciationStartDate_0L_01_Value": "2026-04-01",
                "DepreciationStartDate_0L_01_Change": "Y",
                "ValidityStartDate_0L_01_Value": "",  // ← BLANK for CHANGE
                "DepreciationKey_0L_01_Value": "INDD",
                "DepreciationKey_0L_01_Change": "Y",
                "PlannedUsefulLifeInYears_0L_01_Value": "",
                "PlannedUsefulLifeInYears_0L_01_Change": "",
                "PlannedUsefulLifeInPeriods_0L_01_Value": "",
                "PlannedUsefulLifeInPeriods_0L_01_Change": "",
                "ScrapAmountInCoCodeCrcy_0L_01_Value": "0",
                "ScrapAmountInCoCodeCrcy_0L_01_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_01_Value": "",
                "AcqnProdnCostScrapPercent_0L_01_Change": "",

                ...this._getEmptyLedgerFields("0L_15"),
                ...this._getEmptyLedgerFields("2L_32"),
                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ SCENARIO 2: CREATE new valuation - All 4 ledger areas
         * Use Case: Adding depreciation parameters to newly capitalized asset
         * @private
         */
        _getScenario2_CreateAllAreas() {
            return {
                "SequenceNumber": 2,
                "CompanyCode": "1000",
                "FixedAsset": "000000100002",

                "FixedAssetDescription_Value": "",
                "FixedAssetDescription_Change": "",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "2026-04-01",
                "DepreciationStartDate_Value": "2026-04-01",
                "ValidityStartDate_Value": "2026-04-01",

                // ✅ CREATE Area 0L_01 (Book Depreciation)
                "Ledger_ActionCode_0L_01": "01",
                "Valuation_ActionCode_0L_01": "01",
                "TimeBasedValuation_ActionCode_0L_01": "01",
                "NegativeAmountIsAllowed_0L_01": "true",
                "AssetCapitalizationDate_0L_01_Value": "2026-04-01",
                "DepreciationStartDate_0L_01_Value": "2026-04-01",
                "ValidityStartDate_0L_01_Value": "2026-04-01",  // ← REQUIRED for CREATE
                "DepreciationKey_0L_01_Value": "INSL",
                "DepreciationKey_0L_01_Change": "Y",
                "PlannedUsefulLifeInYears_0L_01_Value": "50",
                "PlannedUsefulLifeInYears_0L_01_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_01_Value": "600",
                "PlannedUsefulLifeInPeriods_0L_01_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_01_Value": "500",
                "ScrapAmountInCoCodeCrcy_0L_01_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_01_Value": "0",
                "AcqnProdnCostScrapPercent_0L_01_Change": "Y",

                // ✅ CREATE Area 0L_15 (Tax Depreciation)
                "Ledger_ActionCode_0L_15": "01",
                "Valuation_ActionCode_0L_15": "01",
                "TimeBasedValuation_ActionCode_0L_15": "01",
                "NegativeAmountIsAllowed_0L_15": "true",
                "AssetCapitalizationDate_0L_15_Value": "2026-04-01",
                "DepreciationStartDate_0L_15_Value": "2026-04-01",
                "ValidityStartDate_0L_15_Value": "2026-04-01",
                "DepreciationKey_0L_15_Value": "INSL",
                "DepreciationKey_0L_15_Change": "Y",
                "PlannedUsefulLifeInYears_0L_15_Value": "55",
                "PlannedUsefulLifeInYears_0L_15_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_15_Value": "660",
                "PlannedUsefulLifeInPeriods_0L_15_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_15_Value": "500",
                "ScrapAmountInCoCodeCrcy_0L_15_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_15_Value": "0",
                "AcqnProdnCostScrapPercent_0L_15_Change": "Y",

                // ✅ CREATE Area 2L_32 (Consolidation) — no Ledger_ActionCode for 2L
                "Valuation_ActionCode_2L_32": "01",
                "TimeBasedValuation_ActionCode_2L_32": "01",
                "NegativeAmountIsAllowed_2L_32": "true",
                "AssetCapitalizationDate_2L_32_Value": "2026-04-01",
                "DepreciationStartDate_2L_32_Value": "2026-04-01",
                "ValidityStartDate_2L_32_Value": "2026-04-01",
                "DepreciationKey_2L_32_Value": "INSL",
                "DepreciationKey_2L_32_Change": "Y",
                "PlannedUsefulLifeInYears_2L_32_Value": "50",
                "PlannedUsefulLifeInYears_2L_32_Change": "Y",
                "PlannedUsefulLifeInPeriods_2L_32_Value": "600",
                "PlannedUsefulLifeInPeriods_2L_32_Change": "Y",
                "ScrapAmountInCoCodeCrcy_2L_32_Value": "500",
                "ScrapAmountInCoCodeCrcy_2L_32_Change": "Y",
                "AcqnProdnCostScrapPercent_2L_32_Value": "0",
                "AcqnProdnCostScrapPercent_2L_32_Change": "Y",

                // ✅ CREATE Area 3L_34 (Special Purpose) — no Ledger_ActionCode for 3L
                "Valuation_ActionCode_3L_34": "01",
                "TimeBasedValuation_ActionCode_3L_34": "01",
                "NegativeAmountIsAllowed_3L_34": "true",
                "AssetCapitalizationDate_3L_34_Value": "2026-04-01",
                "DepreciationStartDate_3L_34_Value": "2026-04-01",
                "ValidityStartDate_3L_34_Value": "2026-04-01",
                "DepreciationKey_3L_34_Value": "INSL",
                "DepreciationKey_3L_34_Change": "Y",
                "PlannedUsefulLifeInYears_3L_34_Value": "50",
                "PlannedUsefulLifeInYears_3L_34_Change": "Y",
                "PlannedUsefulLifeInPeriods_3L_34_Value": "600",
                "PlannedUsefulLifeInPeriods_3L_34_Change": "Y",
                "ScrapAmountInCoCodeCrcy_3L_34_Value": "500",
                "ScrapAmountInCoCodeCrcy_3L_34_Change": "Y",
                "AcqnProdnCostScrapPercent_3L_34_Value": "0",
                "AcqnProdnCostScrapPercent_3L_34_Change": "Y"
            };
        }

        /**
         * ✅ SCENARIO 3: DELETE ledger area
         * Use Case: Removing old depreciation area no longer needed
         * @private
         */
        _getScenario3_DeleteLedger() {
            return {
                "SequenceNumber": 3,
                "CompanyCode": "1000",
                "FixedAsset": "000000100003",

                "FixedAssetDescription_Value": "",
                "FixedAssetDescription_Change": "",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                ...this._getEmptyLedgerFields("0L_01"),
                ...this._getEmptyLedgerFields("0L_15"),

                // ✅ DELETE entire ledger 2L_32 — no Ledger_ActionCode for 2L
                "Valuation_ActionCode_2L_32": "03",
                "AssetCapitalizationDate_2L_32_Value": "",
                "DepreciationStartDate_2L_32_Value": "",
                "ValidityStartDate_2L_32_Value": "",  // ← Not needed for DELETE
                "DepreciationKey_2L_32_Value": "",
                "DepreciationKey_2L_32_Change": "",
                "PlannedUsefulLifeInYears_2L_32_Value": "",
                "PlannedUsefulLifeInYears_2L_32_Change": "",
                "PlannedUsefulLifeInPeriods_2L_32_Value": "",
                "PlannedUsefulLifeInPeriods_2L_32_Change": "",
                "ScrapAmountInCoCodeCrcy_2L_32_Value": "",
                "ScrapAmountInCoCodeCrcy_2L_32_Change": "",
                "AcqnProdnCostScrapPercent_2L_32_Value": "",
                "AcqnProdnCostScrapPercent_2L_32_Change": "",

                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ SCENARIO 4: MIXED operations - CHANGE, CREATE, DELETE in different areas
         * Use Case: Complex multi-area update
         * @private
         */
        _getScenario4_MixedOperations() {
            return {
                "SequenceNumber": 4,
                "CompanyCode": "2000",
                "FixedAsset": "000000200001",

                "FixedAssetDescription_Value": "Multi-Area Asset Update 1",
                "FixedAssetDescription_Change": "Y",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "2026-02-01",
                "DepreciationStartDate_Value": "2026-02-01",
                "ValidityStartDate_Value": "2026-02-01",

                // ✅ CHANGE Area 0L_01 - Update depreciation key
                "Ledger_ActionCode_0L_01": "02",
                "Valuation_ActionCode_0L_01": "02",
                "TimeBasedValuation_ActionCode_0L_01": "02",
                "NegativeAmountIsAllowed_0L_01": "true",
                "AssetCapitalizationDate_0L_01_Value": "",
                "DepreciationStartDate_0L_01_Value": "",
                "ValidityStartDate_0L_01_Value": "",  // ← BLANK for CHANGE
                "DepreciationKey_0L_01_Value": "INDB",
                "DepreciationKey_0L_01_Change": "Y",
                "PlannedUsefulLifeInYears_0L_01_Value": "",
                "PlannedUsefulLifeInYears_0L_01_Change": "",
                "PlannedUsefulLifeInPeriods_0L_01_Value": "",
                "PlannedUsefulLifeInPeriods_0L_01_Change": "",
                "ScrapAmountInCoCodeCrcy_0L_01_Value": "0",
                "ScrapAmountInCoCodeCrcy_0L_01_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_01_Value": "",
                "AcqnProdnCostScrapPercent_0L_01_Change": "",

                // ✅ CREATE Area 0L_15 - Add new tax depreciation
                "Ledger_ActionCode_0L_15": "01",
                "Valuation_ActionCode_0L_15": "01",
                "TimeBasedValuation_ActionCode_0L_15": "01",
                "NegativeAmountIsAllowed_0L_15": "true",
                "AssetCapitalizationDate_0L_15_Value": "2026-02-15",
                "DepreciationStartDate_0L_15_Value": "2026-02-15",
                "ValidityStartDate_0L_15_Value": "2026-02-15",  // ← REQUIRED for CREATE
                "DepreciationKey_0L_15_Value": "INTX",
                "DepreciationKey_0L_15_Change": "Y",
                "PlannedUsefulLifeInYears_0L_15_Value": "10",
                "PlannedUsefulLifeInYears_0L_15_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_15_Value": "120",
                "PlannedUsefulLifeInPeriods_0L_15_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_15_Value": "0",
                "ScrapAmountInCoCodeCrcy_0L_15_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_15_Value": "",
                "AcqnProdnCostScrapPercent_0L_15_Change": "",

                ...this._getEmptyLedgerFields("2L_32"),

                // ✅ DELETE valuation in Area 3L_34 — no Ledger_ActionCode for 3L
                "Valuation_ActionCode_3L_34": "03",
                "AssetCapitalizationDate_3L_34_Value": "",
                "DepreciationStartDate_3L_34_Value": "",
                "ValidityStartDate_3L_34_Value": "",  // ← Not needed for DELETE
                "DepreciationKey_3L_34_Value": "",
                "DepreciationKey_3L_34_Change": "",
                "PlannedUsefulLifeInYears_3L_34_Value": "",
                "PlannedUsefulLifeInYears_3L_34_Change": "",
                "PlannedUsefulLifeInPeriods_3L_34_Value": "",
                "PlannedUsefulLifeInPeriods_3L_34_Change": "",
                "ScrapAmountInCoCodeCrcy_3L_34_Value": "",
                "ScrapAmountInCoCodeCrcy_3L_34_Change": "",
                "AcqnProdnCostScrapPercent_3L_34_Value": "",
                "AcqnProdnCostScrapPercent_3L_34_Change": ""
            };
        }

        /**
         * ✅ SCENARIO 5: General changes only - No depreciation changes
         * Use Case: Updating descriptive and account assignment fields
         * @private
         */
        _getScenario5_GeneralFieldsOnly() {
            return {
                "SequenceNumber": 5,
                "CompanyCode": "3000",
                "FixedAsset": "000000300001",

                "FixedAssetDescription_Value": "Updated Asset Description",
                "FixedAssetDescription_Change": "Y",
                "AssetAdditionalDescription_Value": "Additional Info Added",
                "AssetAdditionalDescription_Change": "Y",
                "AssetSerialNumber_Value": "SN-2026-001",
                "AssetSerialNumber_Change": "Y",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "WBS-GENERAL-UPD",
                "YY1_WBSELEMENT1_FAA_Change": "Y",
                "Room_Value": "C303",
                "Room_Change": "Y",

                "InventoryNote_Value": "Checked and verified",
                "InventoryNote_Change": "Y",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "02",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "CC3000",
                "CostCenter_Change": "Y",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "IN-005",
                "Inventory_Change": "Y",
                "InventoryIsCounted_Value": "N",
                "InventoryIsCounted_Change": "Y",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                // No depreciation changes
                ...this._getEmptyLedgerFields("0L_01"),
                ...this._getEmptyLedgerFields("0L_15"),
                ...this._getEmptyLedgerFields("2L_32"),
                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ SCENARIO 6: Account assignment change
         * Use Case: Moving asset to new cost center and WBS — now includes new AA fields
         * @private
         */
        _getScenario6_AccountAssignmentChange() {
            return {
                "SequenceNumber": 6,
                "CompanyCode": "1000",
                "FixedAsset": "000000100006",

                "FixedAssetDescription_Value": "",
                "FixedAssetDescription_Change": "",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "W-TPWT6RJH2182",
                "WBSElementExternalID_Change": "Y",
                "YY1_WBSELEMENT1_FAA_Value": "W-TPWT6RJH2182",
                "YY1_WBSELEMENT1_FAA_Change": "Y",
                "Room_Value": "B201",
                "Room_Change": "Y",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                // ✅ CHANGE account assignment — full field coverage
                "AccountAssignment_ActionCode": "02",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "2JH_WLA",
                "CostCenter_Change": "Y",
                "Plant_Value": "1001",
                "Plant_Change": "Y",
                "ProfitCenter_Value": "PC1000",
                "ProfitCenter_Change": "Y",
                // NEW v3.0 fields
                "Segment_Value": "SEG01",
                "Segment_Change": "Y",
                "AssetLocation_Value": "LOC01",
                "AssetLocation_Change": "Y",
                "BusinessPlace_Value": "BP01",
                "BusinessPlace_Change": "Y",
                "FunctionalArea_Value": "FA01",
                "FunctionalArea_Change": "Y",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                // No depreciation changes
                ...this._getEmptyLedgerFields("0L_01"),
                ...this._getEmptyLedgerFields("0L_15"),
                ...this._getEmptyLedgerFields("2L_32"),
                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ SCENARIO 7: India-specific asset with R&D designation
         * Use Case: Marking India asset as Research & Development
         * @private
         */
        _getScenario7_IndiaAssetWithRnD() {
            return {
                "SequenceNumber": 7,
                "CompanyCode": "2000",
                "FixedAsset": "000000200007",

                "FixedAssetDescription_Value": "Research Equipment India",
                "FixedAssetDescription_Change": "Y",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                // ✅ India-specific fields populated
                "IN_AssetBlock_Value": "BLK01",
                "IN_AssetBlock_Change": "Y",
                "IN_AssetPutToUseDate_Value": "2026-01-15",
                "IN_AssetPutToUseDate_Change": "Y",
                "IN_AdditionalAssetBlock_Value": "ABLK1",
                "IN_AdditionalAssetBlock_Change": "Y",
                "IN_AssetIsResearchAndDev_Value": "Y",
                "IN_AssetIsResearchAndDev_Change": "Y",
                "IN_AssetIsPriorYear_Value": "N",
                "IN_AssetIsPriorYear_Change": "Y",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                // No depreciation changes
                ...this._getEmptyLedgerFields("0L_01"),
                ...this._getEmptyLedgerFields("0L_15"),
                ...this._getEmptyLedgerFields("2L_32"),
                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ SCENARIO 8: Useful life correction across all areas
         * Use Case: Correcting useful life estimates for depreciation
         * @private
         */
        _getScenario8_UsefulLifeCorrection() {
            return {
                "SequenceNumber": 8,
                "CompanyCode": "1000",
                "FixedAsset": "000000100008",

                "FixedAssetDescription_Value": "",
                "FixedAssetDescription_Change": "",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "",
                "WBSElementExternalID_Change": "",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "",
                "Room_Change": "",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                "IN_AssetBlock_Value": "",
                "IN_AssetBlock_Change": "",
                "IN_AssetPutToUseDate_Value": "",
                "IN_AssetPutToUseDate_Change": "",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                "AccountAssignment_ActionCode": "",
                "AccountAssignment_ValidityStartDate": "",
                "CostCenter_Value": "",
                "CostCenter_Change": "",
                "Plant_Value": "",
                "Plant_Change": "",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                "AssetCapitalizationDate_Value": "",
                "DepreciationStartDate_Value": "",
                "ValidityStartDate_Value": "",

                // ✅ CHANGE useful life in Area 0L_01
                "Ledger_ActionCode_0L_01": "02",
                "Valuation_ActionCode_0L_01": "02",
                "TimeBasedValuation_ActionCode_0L_01": "02",
                "NegativeAmountIsAllowed_0L_01": "true",
                "AssetCapitalizationDate_0L_01_Value": "",
                "DepreciationStartDate_0L_01_Value": "",
                "ValidityStartDate_0L_01_Value": "",  // ← BLANK for CHANGE
                "DepreciationKey_0L_01_Value": "",
                "DepreciationKey_0L_01_Change": "",
                "PlannedUsefulLifeInYears_0L_01_Value": "45",
                "PlannedUsefulLifeInYears_0L_01_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_01_Value": "540",
                "PlannedUsefulLifeInPeriods_0L_01_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_01_Value": "0",
                "ScrapAmountInCoCodeCrcy_0L_01_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_01_Value": "",
                "AcqnProdnCostScrapPercent_0L_01_Change": "",

                // ✅ CHANGE useful life in Area 0L_15
                "Ledger_ActionCode_0L_15": "02",
                "Valuation_ActionCode_0L_15": "02",
                "TimeBasedValuation_ActionCode_0L_15": "02",
                "NegativeAmountIsAllowed_0L_15": "true",
                "AssetCapitalizationDate_0L_15_Value": "",
                "DepreciationStartDate_0L_15_Value": "",
                "ValidityStartDate_0L_15_Value": "",
                "DepreciationKey_0L_15_Value": "",
                "DepreciationKey_0L_15_Change": "",
                "PlannedUsefulLifeInYears_0L_15_Value": "48",
                "PlannedUsefulLifeInYears_0L_15_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_15_Value": "576",
                "PlannedUsefulLifeInPeriods_0L_15_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_15_Value": "0",
                "ScrapAmountInCoCodeCrcy_0L_15_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_15_Value": "",
                "AcqnProdnCostScrapPercent_0L_15_Change": "",

                ...this._getEmptyLedgerFields("2L_32"),
                ...this._getEmptyLedgerFields("3L_34")
            };
        }

        /**
         * ✅ NEW SCENARIO 9: Full payload example
         * Use Case: Mirrors exact SAP OData payload structure from EDMX analysis.
         * Shows: 0L with Ledger_ActionCode, 2L/3L without, independent TimeBasedValuation_ActionCode,
         *        NegativeAmountIsAllowed, full account assignment, India fields.
         * @private
         */
        _getScenario9_FullPayloadExample() {
            return {
                "SequenceNumber": 9,
                "CompanyCode": "2000",
                "FixedAsset": "000010000000",

                "FixedAssetDescription_Value": "Full Payload Example Asset",
                "FixedAssetDescription_Change": "Y",
                "AssetAdditionalDescription_Value": "",
                "AssetAdditionalDescription_Change": "",
                "AssetSerialNumber_Value": "",
                "AssetSerialNumber_Change": "",

                "WBSElementExternalID_Value": "W-TPWT6RJH2182",
                "WBSElementExternalID_Change": "Y",
                "YY1_WBSELEMENT1_FAA_Value": "",
                "YY1_WBSELEMENT1_FAA_Change": "",
                "Room_Value": "12",
                "Room_Change": "Y",

                "InventoryNote_Value": "",
                "InventoryNote_Change": "",

                // India fields
                "IN_AssetBlock_Value": "Z100",
                "IN_AssetBlock_Change": "Y",
                "IN_AssetPutToUseDate_Value": "2010-07-16",
                "IN_AssetPutToUseDate_Change": "Y",
                "IN_AdditionalAssetBlock_Value": "",
                "IN_AdditionalAssetBlock_Change": "",
                "IN_AssetIsResearchAndDev_Value": "",
                "IN_AssetIsResearchAndDev_Change": "",
                "IN_AssetIsPriorYear_Value": "",
                "IN_AssetIsPriorYear_Change": "",

                // Full account assignment
                "AccountAssignment_ActionCode": "01",
                "AccountAssignment_ValidityStartDate": "2026-01-20",
                "CostCenter_Value": "2JH_WLA",
                "CostCenter_Change": "Y",
                "Plant_Value": "1001",
                "Plant_Change": "Y",
                "ProfitCenter_Value": "",
                "ProfitCenter_Change": "",

                "Inventory_Value": "",
                "Inventory_Change": "",
                "InventoryIsCounted_Value": "",
                "InventoryIsCounted_Change": "",

                // Global date fallbacks
                "AssetCapitalizationDate_Value": "2010-07-16",
                "DepreciationStartDate_Value": "2010-07-16",
                "ValidityStartDate_Value": "2026-01-20",

                // ✅ 0L_01: Ledger 0L HAS Ledger_ActionCode, TBV_ActionCode same as Valuation
                "Ledger_ActionCode_0L_01": "02",
                "AssetCapitalizationDate_0L_01_Value": "2010-07-16",
                "AssetCapitalizationDate_0L_01_Change": "Y",
                "Valuation_ActionCode_0L_01": "02",
                "NegativeAmountIsAllowed_0L_01": "true",
                "DepreciationStartDate_0L_01_Value": "2010-07-16",
                "DepreciationStartDate_0L_01_Change": "Y",
                "TimeBasedValuation_ActionCode_0L_01": "02",
                "ValidityStartDate_0L_01_Value": "2026-01-20",
                "DepreciationKey_0L_01_Value": "INDD",
                "DepreciationKey_0L_01_Change": "Y",
                "PlannedUsefulLifeInYears_0L_01_Value": "50",
                "PlannedUsefulLifeInYears_0L_01_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_01_Value": "0",
                "PlannedUsefulLifeInPeriods_0L_01_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_01_Value": "500",
                "ScrapAmountInCoCodeCrcy_0L_01_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_01_Value": "0",
                "AcqnProdnCostScrapPercent_0L_01_Change": "Y",

                // ✅ 0L_15: Ledger 0L HAS Ledger_ActionCode
                "Ledger_ActionCode_0L_15": "02",
                "Valuation_ActionCode_0L_15": "02",
                "NegativeAmountIsAllowed_0L_15": "true",
                "DepreciationStartDate_0L_15_Value": "2010-07-16",
                "DepreciationStartDate_0L_15_Change": "Y",
                "TimeBasedValuation_ActionCode_0L_15": "02",
                "ValidityStartDate_0L_15_Value": "2026-01-20",
                "DepreciationKey_0L_15_Value": "INDD",
                "DepreciationKey_0L_15_Change": "Y",
                "PlannedUsefulLifeInYears_0L_15_Value": "55",
                "PlannedUsefulLifeInYears_0L_15_Change": "Y",
                "PlannedUsefulLifeInPeriods_0L_15_Value": "0",
                "PlannedUsefulLifeInPeriods_0L_15_Change": "Y",
                "ScrapAmountInCoCodeCrcy_0L_15_Value": "200",
                "ScrapAmountInCoCodeCrcy_0L_15_Change": "Y",
                "AcqnProdnCostScrapPercent_0L_15_Value": "0",
                "AcqnProdnCostScrapPercent_0L_15_Change": "Y",

                // ✅ 2L_32: Ledger 2L has NO Ledger_ActionCode
                //    TimeBasedValuation_ActionCode = "01" (DIFFERENT from Valuation "02")
                "Valuation_ActionCode_2L_32": "02",
                "NegativeAmountIsAllowed_2L_32": "true",
                "DepreciationStartDate_2L_32_Value": "2010-07-16",
                "DepreciationStartDate_2L_32_Change": "Y",
                "TimeBasedValuation_ActionCode_2L_32": "01",  // ← INDEPENDENT: CREATE time-based
                "ValidityStartDate_2L_32_Value": "2026-01-20",
                "DepreciationKey_2L_32_Value": "9ADD",
                "DepreciationKey_2L_32_Change": "Y",
                "PlannedUsefulLifeInYears_2L_32_Value": "60",
                "PlannedUsefulLifeInYears_2L_32_Change": "Y",
                "PlannedUsefulLifeInPeriods_2L_32_Value": "0",
                "PlannedUsefulLifeInPeriods_2L_32_Change": "Y",
                "ScrapAmountInCoCodeCrcy_2L_32_Value": "500",
                "ScrapAmountInCoCodeCrcy_2L_32_Change": "Y",
                "AcqnProdnCostScrapPercent_2L_32_Value": "0",
                "AcqnProdnCostScrapPercent_2L_32_Change": "Y",

                // ✅ 3L_34: Ledger 3L has NO Ledger_ActionCode
                //    TimeBasedValuation_ActionCode = "01" (DIFFERENT from Valuation "02")
                "Valuation_ActionCode_3L_34": "02",
                "NegativeAmountIsAllowed_3L_34": "true",
                "DepreciationStartDate_3L_34_Value": "2010-07-16",
                "DepreciationStartDate_3L_34_Change": "Y",
                "TimeBasedValuation_ActionCode_3L_34": "01",  // ← INDEPENDENT: CREATE time-based
                "ValidityStartDate_3L_34_Value": "2026-01-20",
                "DepreciationKey_3L_34_Value": "9NDD",
                "DepreciationKey_3L_34_Change": "Y",
                "PlannedUsefulLifeInYears_3L_34_Value": "75",
                "PlannedUsefulLifeInYears_3L_34_Change": "Y",
                "PlannedUsefulLifeInPeriods_3L_34_Value": "0",
                "PlannedUsefulLifeInPeriods_3L_34_Change": "Y",
                "ScrapAmountInCoCodeCrcy_3L_34_Value": "100",
                "ScrapAmountInCoCodeCrcy_3L_34_Change": "Y",
                "AcqnProdnCostScrapPercent_3L_34_Value": "0",
                "AcqnProdnCostScrapPercent_3L_34_Change": "Y"
            };
        }

        /**
         * Gets all 9 comprehensive example rows
         * @private
         */
        _getAllExampleRows() {
            return [
                this._getScenario1_SimpleChange(),
                this._getScenario2_CreateAllAreas(),
                this._getScenario3_DeleteLedger(),
                this._getScenario4_MixedOperations(),
                this._getScenario5_GeneralFieldsOnly(),
                this._getScenario6_AccountAssignmentChange(),
                this._getScenario7_IndiaAssetWithRnD(),
                this._getScenario8_UsefulLifeCorrection(),
                this._getScenario9_FullPayloadExample()
            ];
        }

        // ─────────────────────────────────────────────────────────────
        // DATA SHEET CREATION
        // ─────────────────────────────────────────────────────────────

        /**
         * Creates the data sheet with headers and example rows
         * @private
         */
        async _createDataSheet(headers, exampleRows) {
            const aoa = [headers];
            exampleRows.forEach(row => {
                const rowData = headers.map(header => {
                    const value = row[header];
                    return value !== undefined && value !== null ? value : "";
                });
                aoa.push(rowData);
            });

            const ws = window.XLSX.utils.aoa_to_sheet(aoa);
            this._applyColumnWidths(ws);
            return ws;
        }

        // ─────────────────────────────────────────────────────────────
        // REFERENCE SHEETS
        // ─────────────────────────────────────────────────────────────

        /**
         * NEW v3.0: Ledger Areas reference sheet
         * @private
         */
        _getLedgerAreasReferenceData() {
            const rows = [
                ["CONFIGURED LEDGER / DEPRECIATION AREAS"],
                [`This template covers ${this.LEDGER_AREAS.length} ledger-area combinations.`],
                ["To add more areas: pass options.ledgerAreas to the AssetMasterChangeTemplateGenerator constructor."],
                [],
                ["Ledger_Area", "Ledger", "Dep.Area", "Ledger_ActionCode column?", "Notes"],
                []
            ];
            for (const la of this.LEDGER_AREAS) {
                const [ledger, area] = la.split("_");
                const hasAC = this.LEDGER_ACTION_CODE_LEDGERS.includes(ledger)
                    ? "YES — has Ledger_ActionCode column"
                    : "NO — no ledger-level FixedAssetObjectActionCode in SAP payload";
                rows.push([la, ledger, area, hasAC, ""]);
            }
            rows.push(
                [],
                ["COLUMN STRUCTURE PER AREA (v3.0)"],
                [],
                ["Column Pattern", "SAP Level", "Entity Type", "Required?", "Notes"],
                ["Ledger_ActionCode_*",               "Ledger",            "FixedAssetLedger_Type",           "If 0L only",          "Only for LEDGER_ACTION_CODE_LEDGERS"],
                ["AssetCapitalizationDate_*_V/C",     "Ledger",            "FixedAssetLedger_Type",           "No",                  ""],
                ["AssetDeactivationDate_*_V/C",       "Ledger",            "FixedAssetLedger_Type",           "No",                  "NEW v3.0"],
                ["Valuation_ActionCode_*",            "Valuation",         "FixedAssetValuation_Type",        "If touching area",    "01/02/03"],
                ["NegativeAmountIsAllowed_*",         "Valuation",         "FixedAssetValuation_Type",        "No",                  "NEW v3.0 — boolean, no _Change pair"],
                ["DepreciationStartDate_*_V/C",       "Valuation",         "FixedAssetValuation_Type",        "No",                  ""],
                ["SpecialDeprStartDate_*_V/C",        "Valuation",         "FixedAssetValuation_Type",        "No",                  "NEW v3.0"],
                ["InterestCalcStartDate_*_V/C",       "Valuation",         "FixedAssetValuation_Type",        "No",                  "NEW v3.0"],
                ["TimeBasedValuation_ActionCode_*",   "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                 "NEW v3.0 — independent from Valuation_ActionCode"],
                ["ValidityStartDate_*_Value",         "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","CREATE=YES CHANGE=NO",""],
                ["DepreciationKey_*_V/C",             "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","CREATE=YES",          ""],
                ["PlannedUsefulLifeInYears_*_V/C",    "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","CREATE=YES",          ""],
                ["PlannedUsefulLifeInPeriods_*_V/C",  "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  ""],
                ["ScrapAmountInCoCodeCrcy_*_V/C",     "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","YES (use 0)",         "FAA_MD/225 if missing"],
                ["AcqnProdnCostScrapPercent_*_V/C",   "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  ""],
                ["VariableDeprPercent_*_V/C",         "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  "NEW v3.0"],
                ["DeprCalcBaseValuePercent_*_V/C",    "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  "NEW v3.0"],
                ["ShiftOperationFactor_*_V/C",        "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  "NEW v3.0"],
                ["IsShutDown_*",                      "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  "NEW v3.0 — boolean, no _Change pair"],
                ["AssetRevaluationIndex_*_V/C",       "TimeBasedValuation","FixedAssetTimeBasedValuation_Type","No",                  "NEW v3.0"]
            );
            return rows;
        }

        /**
         * ✅ Scenarios explanation sheet — updated for 9 scenarios
         * @private
         */
        _getScenariosExplanationData() {
            return [
                ["SCENARIO EXAMPLES GUIDE"],
                ["Each row in the Change_Data sheet demonstrates a different use case"],
                [],
                ["Row #", "Scenario", "Use Case", "Key Features"],
                [],
                ["1", "Simple CHANGE",        "Update depreciation key only",           "• Action Code 02\n• ValidityStartDate blank\n• Only 0L_01 area\n• NegativeAmountIsAllowed + TimeBasedValuation_AC"],
                [],
                ["2", "CREATE All Areas",     "Add depreciation to new asset",          "• Action Code 01\n• ValidityStartDate required\n• All 4 areas\n• 2L/3L: no Ledger_ActionCode"],
                [],
                ["3", "DELETE Ledger",        "Remove old depreciation area",           "• Action Code 03 on Valuation_2L_32\n• 2L has no Ledger_ActionCode\n• No dates needed"],
                [],
                ["4", "MIXED Operations",     "Complex multi-area update",              "• CHANGE 0L_01\n• CREATE 0L_15\n• DELETE 3L_34 valuation\n• Shows all 3 actions"],
                [],
                ["5", "General Fields Only",  "Update descriptive info",                "• No depreciation changes\n• Description, serial, room\n• Account assignment, inventory"],
                [],
                ["6", "Account Assignment",   "Move asset to new cost center",          "• Full AA fields: Segment, AssetLocation, BusinessPlace, FunctionalArea\n• ActionCode 02"],
                [],
                ["7", "India R&D Asset",      "Mark India asset as R&D",               "• IN_AssetBlock, IN_AssetIsResearchAndDev = Y\n• IN_AdditionalAssetBlock\n• Complete India fields"],
                [],
                ["8", "Useful Life Correction","Fix depreciation estimates",            "• Update years/periods in 0L_01 and 0L_15\n• ScrapAmount required\n• TimeBasedValuation_AC = 02"],
                [],
                ["9", "Full Payload Example", "Exact SAP OData payload structure",      "• ALL new v3.0 fields\n• 0L: has Ledger_ActionCode\n• 2L/3L: NO Ledger_ActionCode\n• TBV_ActionCode = 01 for 2L/3L (independent)\n• NegativeAmountIsAllowed per area"],
                [],
                ["HOW TO USE THESE EXAMPLES"],
                [],
                ["1. Find the scenario that matches your need"],
                ["2. Copy the entire row from Change_Data sheet"],
                ["3. Paste into a new row"],
                ["4. Modify the values for your asset"],
                ["5. Keep the same pattern (Action Codes, ValidityStartDate rules, etc.)"],
                [],
                ["CRITICAL REMINDERS"],
                [],
                ["✓ CREATE (01): ValidityStartDate REQUIRED with NEW date"],
                ["✓ CHANGE (02): ValidityStartDate LEAVE BLANK (recommended)"],
                ["✓ DELETE (03): ValidityStartDate NOT NEEDED"],
                ["✓ ScrapAmountInCoCodeCrcy: ALWAYS provide (use 0 if no scrap)"],
                ["✓ _Change flags: ALWAYS set to Y when providing _Value"],
                ["✓ NegativeAmountIsAllowed: true/false — no _Change column needed"],
                ["✓ TimeBasedValuation_ActionCode: independent from Valuation_ActionCode"],
                ["✓ Ledger_ActionCode: ONLY for 0L areas — 2L and 3L do NOT have this column"],
                [],
                ["QUICK REFERENCE"],
                [],
                ["Need to...", "Use Scenario #"],
                ["Just change depreciation key", "1"],
                ["Set up new asset depreciation", "2"],
                ["Remove old depreciation area", "3"],
                ["Do multiple different changes", "4"],
                ["Update only descriptions/info", "5"],
                ["Move asset to new location/cost center", "6"],
                ["Mark India asset as R&D", "7"],
                ["Fix useful life estimates", "8"],
                ["Mirror exact SAP payload structure", "9"]
            ];
        }

        /**
         * ✅ Instructions data — updated for v3.0
         * @private
         */
        _getInstructionsData() {
            return [
                ["ASSET MASTER CHANGE TEMPLATE - INSTRUCTIONS"],
                ["Version 3.0 - Dynamic Ledger Areas + Full EDMX Field Coverage (9 Scenarios)"],
                [],
                ["================== OVERVIEW =================="],
                ["This template is used to perform mass changes to Fixed Asset Master data in SAP."],
                ["You can Create, Change, or Delete asset master data using this template."],
                [],
                ["================== NEW IN V3.0 =================="],
                ["• Dynamic ledger areas — configure constructor with any chart of depreciation"],
                ["• NegativeAmountIsAllowed per area (FixedAssetValuation_Type, boolean, no _Change)"],
                ["• TimeBasedValuation_ActionCode per area — independent from Valuation_ActionCode"],
                ["• Ledger_ActionCode ONLY for 0L ledger — 2L and 3L do NOT have this column"],
                ["• Full FixedAssetAssignment_Type: Segment, AssetLocation, Fund, GrantID, BusinessPlace, etc."],
                ["• Additional date fields: SpecialDeprStartDate, InterestCalcStartDate, AssetDeactivationDate"],
                ["• Additional valuation fields: VariableDeprPercent, ShiftOperationFactor, IsShutDown, etc."],
                ["• Scenario 9: Full payload example mirroring exact SAP OData structure"],
                ["• New 'Ledger_Areas' reference sheet"],
                [],
                ["================== 9 SAMPLE SCENARIOS =================="],
                ["The Change_Data sheet includes 9 comprehensive example scenarios:"],
                [""],
                ["Row 1: Simple CHANGE - Update depreciation key (0L_01 only)"],
                ["Row 2: CREATE All Areas - Add depreciation to all 4 areas"],
                ["Row 3: DELETE Ledger - Remove old depreciation area (2L_32)"],
                ["Row 4: MIXED Operations - CHANGE/CREATE/DELETE in one asset"],
                ["Row 5: General Fields Only - Update descriptions/account assignment"],
                ["Row 6: Account Assignment - Move asset (incl. Segment, AssetLocation, etc.)"],
                ["Row 7: India R&D Asset - Mark India asset as Research & Development"],
                ["Row 8: Useful Life Correction - Fix depreciation estimates"],
                ["Row 9: Full Payload Example - Exact SAP OData structure with all new fields"],
                [""],
                ["See the 'Scenarios' sheet for detailed explanation of each example!"],
                [],
                ["================== CHANGE OPERATION CONSTRAINTS =================="],
                ["Important constraints when changing fixed assets:"],
                ["- You cannot change the values of _GlobTimeBasedMasterData and _GlobTimeBasedValuation properties"],
                ["- Account assignment entries cannot be deleted (only created or changed)"],
                ["- Adjacent account assignment entries with same data will be automatically merged"],
                [],
                ["⚠️ CRITICAL: MANDATORY PAIRED FIELDS - SAP BUSINESS RULE ⚠️"],
                ["CompanyCodeCurrency and ScrapAmountInCoCodeCrcy MUST ALWAYS BE PROVIDED TOGETHER"],
                [""],
                ["- CompanyCodeCurrency is automatically set to 'INR' by the system"],
                ["- ScrapAmountInCoCodeCrcy MUST be provided when creating or changing depreciation"],
                ["- If your asset has no scrap value, use 0 (zero)"],
                ["- Omitting ScrapAmountInCoCodeCrcy will cause SAP error: FAA_MD/225"],
                ["- Error message: 'Values must be specified for both fields: for ScrapAmountInCoCodeCrcy and CompanyCodeCurrency'"],
                [],
                ["⚠️ CRITICAL: VALIDITYSTARTDATE REQUIREMENTS (SAP API Cardinality 1..1) ⚠️"],
                ["ValidityStartDate is MANDATORY in SAP for time-based valuation operations."],
                [""],
                ["FOR CREATE OPERATIONS (Action Code 01):"],
                ["✅ ValidityStartDate is MANDATORY - You MUST provide a NEW date"],
                ["   • Use a date that doesn't exist in SAP yet (e.g., 2026-01-20)"],
                ["   • Format: YYYY-MM-DD"],
                ["   • Example: ValidityStartDate_0L_01_Value = 2026-01-20"],
                ["   • This marks when your new depreciation parameters take effect"],
                ["   • See Row 2 in Change_Data sheet for example"],
                [""],
                ["FOR CHANGE OPERATIONS (Action Code 02):"],
                ["⚠️ ValidityStartDate is OPTIONAL - RECOMMENDED TO LEAVE BLANK"],
                ["   • If blank: System automatically uses existing date from SAP (SAFEST)"],
                ["   • If provided: Must EXACTLY match an existing interval in SAP"],
                ["   • Wrong date → ERROR: 'time interval doesn't exist'"],
                ["   • Best practice: Leave ValidityStartDate blank for CHANGE operations"],
                ["   • See Row 1 in Change_Data sheet for example"],
                [""],
                ["EXAMPLES:"],
                ["✓ CREATE: ValidityStartDate_0L_01_Value = 2026-01-20 (NEW date)"],
                ["✓ CHANGE: ValidityStartDate_0L_01_Value = <BLANK> (recommended)"],
                ["✗ WRONG: CREATE with blank ValidityStartDate → ERROR"],
                [],
                ["⚠️ NEW v3.0: LEDGER_ACTIONCODE RULES ⚠️"],
                ["Ledger_ActionCode only applies to Ledger 0L:"],
                ["  ✅ 0L_01, 0L_15: HAVE Ledger_ActionCode column"],
                ["  ❌ 2L_32, 3L_34: DO NOT have Ledger_ActionCode column (SAP payload structure)"],
                ["See Row 9 in Change_Data for example of both cases."],
                [],
                ["⚠️ NEW v3.0: TIMEBASEDVALUATION_ACTIONCODE ⚠️"],
                ["TimeBasedValuation_ActionCode is INDEPENDENT from Valuation_ActionCode."],
                ["Example: Valuation=02 (CHANGE) + TimeBasedValuation=01 (CREATE) is valid."],
                ["See Row 9 for 2L_32 and 3L_34 which use exactly this pattern."],
                [],
                ["================== CRITICAL INFORMATION - LEDGER_AREA FORMAT =================="],
                ["Depreciation areas use Ledger_Area format for better clarity:"],
                ["- 0L_01: Book depreciation (Ledger 0L, Area 01)"],
                ["- 0L_15: Tax depreciation (Ledger 0L, Area 15)"],
                ["- 2L_32: Consolidation depreciation (Ledger 2L, Area 32)"],
                ["- 3L_34: Special purpose depreciation (Ledger 3L, Area 34)"],
                [],
                ["================== DATE HIERARCHY (ENHANCED) =================="],
                ["The system supports both general and ledger-specific dates with fallback hierarchy:"],
                [""],
                ["1. LEDGER-SPECIFIC DATES (Highest Priority):"],
                ["   - AssetCapitalizationDate_0L_01_Value"],
                ["   - DepreciationStartDate_0L_01_Value"],
                ["   - ValidityStartDate_0L_01_Value"],
                ["   (Similar for 0L_15, 2L_32, 3L_34)"],
                [""],
                ["2. GENERAL DATES (Fallback):"],
                ["   - AssetCapitalizationDate_Value"],
                ["   - DepreciationStartDate_Value"],
                ["   - ValidityStartDate_Value"],
                [""],
                ["3. FINAL FALLBACK:"],
                ["   - IN_AssetPutToUseDate_Value"],
                [],
                ["================== ACTION CODES =================="],
                ["ACTION CODE | STRUCTURE          | DESCRIPTION"],
                ["01          | CREATE             | Create new entry - ValidityStartDate MANDATORY"],
                ["02          | CHANGE             | Change existing entry (DEFAULT) - ValidityStartDate OPTIONAL (leave blank)"],
                ["03          | DELETE             | Delete existing entry - ValidityStartDate NOT NEEDED"],
                [],
                ["================== HOW TO USE THIS TEMPLATE =================="],
                [],
                ["STEP 1: CHOOSE YOUR SCENARIO"],
                ["- Look at the 9 example rows in Change_Data sheet"],
                ["- Find the scenario that matches your need"],
                ["- Check the 'Scenarios' sheet for detailed explanations"],
                [],
                ["STEP 2: COPY AND MODIFY"],
                ["- Copy the entire row of your chosen scenario"],
                ["- Paste into a new row"],
                ["- Change CompanyCode, FixedAsset to your asset"],
                ["- Modify the values but keep the same pattern"],
                [],
                ["STEP 3: UNDERSTAND THE PATTERN"],
                ["For each field you want to change:"],
                ["- Enter the new value in the '_Value' column"],
                ["- Mark 'Y', 'YES', 'TRUE', '1', or 'X' in the '_Change' column"],
                ["- To clear a field, enter <BLANK> in the '_Value' column and mark the '_Change' column"],
                ["- Boolean fields (NegativeAmountIsAllowed, IsShutDown): use true/false, no _Change needed"],
                [],
                ["STEP 4: VALIDATE BEFORE UPLOAD"],
                ["⚠️ CRITICAL: Check these before upload:"],
                ["- CREATE (01): ValidityStartDate filled with NEW date"],
                ["- CHANGE (02): ValidityStartDate left BLANK"],
                ["- ScrapAmountInCoCodeCrcy: Always provided (use 0 if no scrap)"],
                ["- _Change flags: Set to Y for every _Value you filled"],
                ["- Ledger_ActionCode: Only fill for 0L areas — leave blank for 2L, 3L"],
                [],
                ["================== COMMON ERRORS & SOLUTIONS =================="],
                [],
                ["❌ ERROR: 'ValidityStartDate is required'"],
                ["   CAUSE: Using CREATE (01) without ValidityStartDate"],
                ["   SOLUTION: Provide ValidityStartDate_0L_01_Value = 2026-01-20"],
                ["   EXAMPLE: See Row 2 in Change_Data"],
                [""],
                ["❌ ERROR: 'Time interval starting with YYYY-MM-DD doesn't exist'"],
                ["   CAUSE: Using CHANGE (02) with wrong ValidityStartDate"],
                ["   SOLUTION: Leave ValidityStartDate blank for CHANGE operations"],
                ["   EXAMPLE: See Row 1 in Change_Data"],
                [""],
                ["❌ ERROR: 'Invalid date format'"],
                ["   CAUSE: Wrong date format (e.g., 01/20/2026)"],
                ["   SOLUTION: Always use YYYY-MM-DD format (e.g., 2026-01-20)"],
                [""],
                ["❌ ERROR: 'Values must be specified for both fields: ScrapAmountInCoCodeCrcy and CompanyCodeCurrency'"],
                ["   CAUSE: Missing ScrapAmountInCoCodeCrcy when changing depreciation"],
                ["   SOLUTION: Always provide ScrapAmountInCoCodeCrcy (use 0 if no scrap value)"],
                ["   EXAMPLE: See all examples in Change_Data - ScrapAmount always provided"],
                [],
                ["================== BEST PRACTICES =================="],
                [],
                ["1. Start with the examples - Don't create from scratch"],
                ["2. For CREATE: Provide ValidityStartDate with a NEW future date"],
                ["3. For CHANGE: Leave ValidityStartDate BLANK (safest approach)"],
                ["4. ALWAYS provide ScrapAmountInCoCodeCrcy when changing depreciation (use 0 if no scrap)"],
                ["5. Test with a small batch first (1-2 assets)"],
                ["6. Validate all dates are in YYYY-MM-DD format"],
                ["7. Review validation warnings before submitting"],
                ["8. Refer to the Scenarios and Ledger_Areas sheets for detailed guidance"],
                [],
                ["For additional help, refer to the 'Action_Codes', 'Field_Reference', 'Ledger_Areas', and 'Scenarios' sheets."]
            ];
        }

        /**
         * ✅ Action codes reference — updated for v3.0
         * @private
         */
        _getActionCodesReferenceData() {
            return [
                ["ACTION CODE REFERENCE - LEDGER_AREA FORMAT WITH VALIDITYSTARTDATE REQUIREMENTS"],
                [],
                ["⚠️ CRITICAL: VALIDITYSTARTDATE VALIDATION RULES ⚠️"],
                ["Based on SAP API documentation (Cardinality 1..1 = MANDATORY field)"],
                [""],
                ["Code", "Action", "ValidityStartDate Requirement", "What Happens", "See Example Row"],
                [],
                ["01", "CREATE", "✅ MANDATORY - Provide NEW date", "Creates new time interval from specified date", "Row 2"],
                ["02", "CHANGE", "⚠️ OPTIONAL - Leave BLANK (recommended)", "Modifies existing interval using existing date", "Row 1, 8"],
                ["02", "CHANGE", "⚠️ ADVANCED - Exact match if provided", "Modifies specific interval if date matches exactly", "(Not recommended)"],
                ["03", "DELETE", "❌ NOT NEEDED - Always ignored", "Deletes entire area, dates not processed", "Row 3"],
                [],
                ["NEW v3.0: ACTION CODE LEVELS"],
                [],
                ["Level",             "Column",                              "Applies to",                    "Notes"],
                ["Ledger",            "Ledger_ActionCode_*",                 "0L areas ONLY",                 "2L/3L do NOT have this column"],
                ["Valuation",         "Valuation_ActionCode_*",              "All areas",                     "Controls _Valuation entry"],
                ["TimeBasedValuation","TimeBasedValuation_ActionCode_*",     "All areas",                     "INDEPENDENT from Valuation_ActionCode"],
                [],
                ["EXAMPLE: Valuation=02 (CHANGE) + TimeBasedValuation=01 (CREATE) is valid."],
                ["See Row 9 for 2L_32 and 3L_34 which use exactly this pattern."],
                [],
                ["═══════════════════════════════════════════════════════════════"],
                ["CREATE (01) - ValidityStartDate is MANDATORY"],
                ["═══════════════════════════════════════════════════════════════"],
                [""],
                ["Rule: You MUST provide a NEW ValidityStartDate that doesn't exist in SAP"],
                [""],
                ["✓ CORRECT EXAMPLE (See Row 2 in Change_Data):"],
                ["  Valuation_ActionCode_0L_01 = 01"],
                ["  TimeBasedValuation_ActionCode_0L_01 = 01"],
                ["  ValidityStartDate_0L_01_Value = 2026-04-01    ← NEW date"],
                ["  DepreciationKey_0L_01_Value = INSL"],
                ["  DepreciationKey_0L_01_Change = Y"],
                ["  PlannedUsefulLifeInYears_0L_01_Value = 50"],
                ["  PlannedUsefulLifeInYears_0L_01_Change = Y"],
                ["  ScrapAmountInCoCodeCrcy_0L_01_Value = 500"],
                ["  ScrapAmountInCoCodeCrcy_0L_01_Change = Y"],
                ["  NegativeAmountIsAllowed_0L_01 = true"],
                [""],
                ["═══════════════════════════════════════════════════════════════"],
                ["CHANGE (02) - ValidityStartDate is OPTIONAL (Leave Blank Recommended)"],
                ["═══════════════════════════════════════════════════════════════"],
                [""],
                ["✓ CORRECT - RECOMMENDED APPROACH (See Row 1 in Change_Data):"],
                ["  Valuation_ActionCode_0L_01 = 02"],
                ["  TimeBasedValuation_ActionCode_0L_01 = 02"],
                ["  ValidityStartDate_0L_01_Value = <BLANK>       ← Recommended!"],
                ["  DepreciationKey_0L_01_Value = INDD"],
                ["  DepreciationKey_0L_01_Change = Y"],
                ["  ScrapAmountInCoCodeCrcy_0L_01_Value = 0"],
                ["  ScrapAmountInCoCodeCrcy_0L_01_Change = Y"],
                ["  NegativeAmountIsAllowed_0L_01 = true"],
                [""],
                ["═══════════════════════════════════════════════════════════════"],
                ["DELETE (03) - ValidityStartDate NOT NEEDED"],
                ["═══════════════════════════════════════════════════════════════"],
                [""],
                ["✓ CORRECT (See Row 3 in Change_Data):"],
                ["  Valuation_ActionCode_2L_32 = 03"],
                ["  (Note: 2L_32 has no Ledger_ActionCode column at all)"],
                ["  (All other fields can be blank, including ValidityStartDate)"],
                [""],
                ["IMPORTANT NOTES:"],
                ["- If no action code specified, system defaults to '02' (CHANGE)"],
                ["- Action codes apply to specific Ledger_Areas (0L_01, 0L_15, 2L_32, 3L_34)"],
                ["- ValidityStartDate requirement applies even when using date fallback hierarchy"],
                ["- 0L Ledger: has Ledger_ActionCode | 2L, 3L Ledgers: NO Ledger_ActionCode"],
                [""],
                ["DATE FORMAT REQUIREMENTS:"],
                ["- Format: YYYY-MM-DD"],
                ["- ✓ Valid: 2026-01-20, 2010-07-16"],
                ["- ✗ Invalid: 01/20/2026, 20-01-2026, 2026.01.20"],
                [""],
                ["COMMON SCENARIOS (SEE CHANGE_DATA EXAMPLES):"],
                [],
                ["Scenario", "Action", "ValidityStartDate", "Example Row"],
                ["Create new depreciation", "CREATE (01)", "2026-01-20 (NEW date)", "Row 2"],
                ["Fix depreciation key", "CHANGE (02)", "<BLANK> (recommended)", "Row 1"],
                ["Remove depreciation area", "DELETE (03)", "<NOT NEEDED>", "Row 3"],
                ["Update useful life", "CHANGE (02)", "<BLANK> (recommended)", "Row 8"],
                ["Full payload structure", "All types", "Various", "Row 9"],
                [""],
                ["⚠️ ScrapAmountInCoCodeCrcy is also MANDATORY for CREATE and CHANGE"],
                ["See 'Field_Reference' sheet for complete requirements"]
            ];
        }

        /**
         * ✅ Field reference — updated for v3.0 new fields
         * @private
         */
        _getFieldReferenceData() {
            return [
                ["FIELD REFERENCE GUIDE - ALL SUPPORTED FIELDS (v3.0)"],
                [],
                ["Field Name", "Entity Type", "SAP Type", "Required", "Notes"],
                [],
                ["── IDENTIFICATION ──"],
                ["SequenceNumber", "N/A",             "String",  "Yes", "Unique identifier for each row"],
                ["CompanyCode",    "N/A",             "String",  "Yes", "4-char company code e.g. 1000"],
                ["FixedAsset",     "N/A",             "String",  "Yes", "12-digit asset number with leading zeros"],
                [],
                ["── GENERAL (FixedAsset_Type) ──"],
                ["FixedAssetDescription",      "FixedAsset_Type", "String(50)", "No",  "Main asset description"],
                ["AssetAdditionalDescription", "FixedAsset_Type", "String(50)", "No",  "Additional description"],
                ["AssetSerialNumber",           "FixedAsset_Type", "String(18)", "No",  "Serial number"],
                [],
                ["── INVENTORY ──"],
                ["Inventory",         "FixedAsset_Type", "String",   "No", "Inventory number"],
                ["InventoryNote",      "FixedAsset_Type", "String(50)","No", "Inventory remarks"],
                ["InventoryIsCounted", "FixedAsset_Type", "Boolean",  "No", "Y/N or true/false"],
                [],
                ["── INDIA-SPECIFIC (FixedAssetCountryData_Type) ──"],
                ["IN_AssetBlock",           "CountryData", "String(5)", "No", "India asset block"],
                ["IN_AssetPutToUseDate",    "CountryData", "Date",      "No", "YYYY-MM-DD"],
                ["IN_AdditionalAssetBlock", "CountryData", "String(5)", "No", "Additional block"],
                ["IN_AssetIsResearchAndDev","CountryData", "String",    "No", "Y/N"],
                ["IN_AssetIsPriorYear",     "CountryData", "String",    "No", "Y/N"],
                [],
                ["── ACCOUNT ASSIGNMENT (FixedAssetAssignment_Type) ──"],
                ["AccountAssignment_ActionCode",       "Assignment", "String", "No",  "01/02/03"],
                ["AccountAssignment_ValidityStartDate","Assignment", "Date",   "No",  "YYYY-MM-DD"],
                ["CostCenter",               "Assignment", "String", "No",  ""],
                ["Plant",                    "Assignment", "String", "No",  ""],
                ["ProfitCenter",             "Assignment", "String", "No",  ""],
                ["Segment",                  "Assignment", "String", "No",  "NEW v3.0"],
                ["WBSElementExternalID",     "Assignment", "String", "No",  ""],
                ["YY1_WBSELEMENT1_FAA",      "Assignment", "String", "No",  "Custom WBS field"],
                ["AssetLocation",            "Assignment", "String", "No",  "NEW v3.0"],
                ["Room",                     "Assignment", "String", "No",  ""],
                ["VehicleLicensePlateNumber","Assignment", "String", "No",  "NEW v3.0"],
                ["TaxJurisdiction",          "Assignment", "String", "No",  "NEW v3.0"],
                ["BusinessPlace",            "Assignment", "String", "No",  "NEW v3.0"],
                ["Fund",                     "Assignment", "String", "No",  "NEW v3.0"],
                ["GrantID",                  "Assignment", "String", "No",  "NEW v3.0"],
                ["FunctionalArea",           "Assignment", "String", "No",  "NEW v3.0"],
                [],
                ["── PER-AREA FIELDS (FixedAssetLedger_Type) ──"],
                ["Ledger_ActionCode_*",           "FixedAssetLedger_Type",  "String",  "0L only",    "ONLY for 0L ledger — 2L/3L have NO Ledger_ActionCode"],
                ["AssetCapitalizationDate_*_V/C", "FixedAssetLedger_Type",  "Date",    "No",          "Ledger-level cap date"],
                ["AssetDeactivationDate_*_V/C",   "FixedAssetLedger_Type",  "Date",    "No",          "NEW v3.0"],
                [],
                ["── PER-AREA FIELDS (FixedAssetValuation_Type) ──"],
                ["Valuation_ActionCode_*",         "FixedAssetValuation_Type", "String",  "Conditional", "01/02/03"],
                ["NegativeAmountIsAllowed_*",      "FixedAssetValuation_Type", "Boolean", "No",          "NEW v3.0 — true/false, no _Change pair"],
                ["DepreciationStartDate_*_V/C",    "FixedAssetValuation_Type", "Date",    "No",          ""],
                ["SpecialDeprStartDate_*_V/C",     "FixedAssetValuation_Type", "Date",    "No",          "NEW v3.0"],
                ["InterestCalcStartDate_*_V/C",    "FixedAssetValuation_Type", "Date",    "No",          "NEW v3.0"],
                [],
                ["── PER-AREA FIELDS (FixedAssetTimeBasedValuation_Type) ──"],
                ["TimeBasedValuation_ActionCode_*",  "TimeBasedValuation_Type", "String",  "No",          "NEW v3.0 — independent from Valuation_ActionCode"],
                ["ValidityStartDate_*_Value",        "TimeBasedValuation_Type", "Date",    "CREATE=YES",  "Mandatory for CREATE, leave blank for CHANGE"],
                ["DepreciationKey_*_V/C",            "TimeBasedValuation_Type", "String(4)","CREATE=YES",  ""],
                ["PlannedUsefulLifeInYears_*_V/C",   "TimeBasedValuation_Type", "String",  "CREATE=YES",  ""],
                ["PlannedUsefulLifeInPeriods_*_V/C", "TimeBasedValuation_Type", "String",  "No",          ""],
                ["ScrapAmountInCoCodeCrcy_*_V/C",    "TimeBasedValuation_Type", "Decimal", "YES (use 0)", "MANDATORY — FAA_MD/225 if missing"],
                ["AcqnProdnCostScrapPercent_*_V/C",  "TimeBasedValuation_Type", "Decimal", "No",          ""],
                ["VariableDeprPercent_*_V/C",        "TimeBasedValuation_Type", "Decimal", "No",          "NEW v3.0"],
                ["DeprCalcBaseValuePercent_*_V/C",   "TimeBasedValuation_Type", "Decimal", "No",          "NEW v3.0"],
                ["ShiftOperationFactor_*_V/C",       "TimeBasedValuation_Type", "Decimal", "No",          "NEW v3.0"],
                ["IsShutDown_*",                     "TimeBasedValuation_Type", "Boolean", "No",          "NEW v3.0 — true/false, no _Change pair"],
                ["AssetRevaluationIndex_*_V/C",      "TimeBasedValuation_Type", "String",  "No",          "NEW v3.0"],
                [],
                ["⚠️ CRITICAL VALIDATION RULES"],
                ["Rule", "CREATE (01)", "CHANGE (02)", "DELETE (03)", "Error if Violated", "Example Row"],
                ["ValidityStartDate", "MANDATORY - provide NEW date", "OPTIONAL - leave blank", "Not needed", "'ValidityStartDate is required for CREATE'", "Row 2 vs Row 1"],
                ["ValidityStartDate format", "Must be YYYY-MM-DD", "Must be YYYY-MM-DD if provided", "N/A", "'Invalid date format'", "All examples"],
                ["ScrapAmountInCoCodeCrcy", "MANDATORY (can be 0)", "MANDATORY (can be 0)", "Not needed", "FAA_MD/225", "All depr examples"],
                ["Ledger_ActionCode", "0L areas only", "0L areas only", "0L areas only", "Wrong payload structure", "Row 9 shows 2L/3L without it"],
                [],
                ["LEDGER_AREA MAPPING"],
                ["Ledger_Area", "Ledger", "Area", "Ledger_ActionCode?", "Description"],
                ["0L_01", "0L", "01", "YES", "Book Depreciation"],
                ["0L_15", "0L", "15", "YES", "Tax Depreciation"],
                ["2L_32", "2L", "32", "NO",  "Consolidation (IFRS)"],
                ["3L_34", "3L", "34", "NO",  "Special Purpose"],
                [],
                ["NOTES:"],
                ["- ValidityStartDate is MANDATORY field in SAP (Cardinality 1..1)"],
                ["- For CREATE: Must provide NEW date that doesn't exist in SAP (see Row 2)"],
                ["- For CHANGE: Recommended to leave blank - system uses existing date (see Row 1)"],
                ["- Always use YYYY-MM-DD format for dates"],
                ["- ScrapAmountInCoCodeCrcy is also MANDATORY for CREATE/CHANGE operations"],
                ["- All 9 example rows in Change_Data sheet follow these rules correctly"],
                ["- Ledger_ActionCode only applies to 0L ledger (see Ledger_Areas sheet)"]
            ];
        }

        // ─────────────────────────────────────────────────────────────
        // STYLING
        // ─────────────────────────────────────────────────────────────

        /**
         * Applies basic styling to the Instructions sheet
         * @private
         */
        _styleInstructionsSheet(ws) {
            if (!ws['!merges']) ws['!merges'] = [];
            if (!ws['!cols']) ws['!cols'] = [];

            ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });

            if (ws['A1']) ws['A1'].s = { font: { sz: 14, bold: true, color: { rgb: "FF000080" } }, alignment: { horizontal: "center" } };
            if (ws['A2']) ws['A2'].s = { font: { sz: 12, bold: true, color: { rgb: "FFFF0000" } } };

            ws['!cols'] = [
                { wch: 45 },
                { wch: 15 },
                { wch: 60 }
            ];
        }

        /**
         * Applies basic styling to reference sheets
         * @private
         */
        _styleReferenceSheet(ws) {
            if (!ws['!cols']) ws['!cols'] = [];

            const range = window.XLSX.utils.decode_range(ws['!ref']);
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = ws[window.XLSX.utils.encode_cell({ r: range.s.r + 2, c: C })];
                if (cell) cell.s = { font: { bold: true }, fill: { fgColor: { rgb: "FFE0E0E0" } } };
            }

            ws['!cols'] = [
                { wch: 40 },
                { wch: 25 },
                { wch: 15 },
                { wch: 20 },
                { wch: 55 }
            ];
        }

        /**
         * Applies default column widths to a worksheet
         * @private
         */
        _applyColumnWidths(ws) {
            const headers = this._getAssetMasterChangeHeaders();
            const colWidths = headers.map(header => {
                if (header.includes("Description"))                               return { wch: 30 };
                if (header.includes("SerialNumber") || header.includes("InventoryNote") || header.includes("WBSElement")) return { wch: 25 };
                if (header.includes("Date"))                                      return { wch: 15 };
                if (header.includes("_Change") || header.includes("IsShutDown") || header.includes("NegativeAmount")) return { wch: 12 };
                if (header.includes("ActionCode"))                                return { wch: 22 };
                if (header.includes("_0L_") || header.includes("_2L_") || header.includes("_3L_")) return { wch: 20 };
                return { wch: 15 };
            });
            ws['!cols'] = colWidths;
        }
    };
});