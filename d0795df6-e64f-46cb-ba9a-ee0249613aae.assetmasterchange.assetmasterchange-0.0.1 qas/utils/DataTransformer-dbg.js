// Fixed DataTransformer.js - Keeps paste1's payload format but fixes the issue
// The problem was in _addLedgerFieldToPayloadWithLedgerArea method - missing debugger and proper _Valuation structure

sap.ui.define([
    "sap/ui/base/Object"
], function (BaseObject) {
    "use strict";

    return BaseObject.extend("assetmasterchange.utils.DataTransformer", {

        constructor: function () {
            BaseObject.call(this);

            // Core configuration with updated Ledger_Area mapping
            this.config = {
                defaultCurrency: "INR",
                ledgerAreaToLedgerMap: {
                    '0L_01': { ledger: '0L', area: '01' },
                    '0L_15': { ledger: '0L', area: '15' },
                    '2L_32': { ledger: '2L', area: '32' },
                    '3L_34': { ledger: '3L', area: '34' }
                },
                areaToLedgerMap: {
                    '01': '0L', '15': '0L', '32': '2L', '34': '3L'
                },
                // v3.0 NEW: only 0L carries Ledger_ActionCode at ledger level per EDMX
                ledgerActionCodeLedgers: ["0L"],
                blankMarkers: ["<BLANK>", "<CLEAR>", "<EMPTY>", "<<BLANK>>"],
                changeFlagValues: {
                    positive: ["Y", "YES", "TRUE", "1", "X", true],
                    negative: ["N", "NO", "FALSE", "0", "", false, null, undefined]
                },
                actionCodes: {
                    create: "01",
                    change: "02",
                    delete: "03"
                },
                validActionCodes: ["01", "02", "03"],
                defaultActionCode: "02",
                validActionsByLevel: {
                    accountAssignment: ["01", "02"],
                    ledger: ["01", "02", "03"],
                    valuation: ["01", "02", "03"],
                    timeBasedValuation: ["01", "02"]
                }
            };
            this.validationConfig = {
                // Mandatory fields by section
                mandatoryFields: {
                    _Root: {
                        CompanyCode: { required: true, maxLength: 4 },
                        FixedAsset: { required: true, maxLength: 16 }
                    },
                    _AccountAssignment: {
                        FixedAssetObjectActionCode: { required: true, values: ["01", "02"] },
                        ValidityStartDate: { required: true, type: "date" } // Can be defaulted
                    },
                    _Ledger: {
                        Ledger: { required: true, maxLength: 2 },
                        FixedAssetObjectActionCode: { required: false, values: ["01", "02", "03"] }
                    },
                    _Valuation: {
                        AssetDepreciationArea: { required: true, maxLength: 2 },
                        FixedAssetObjectActionCode: { required: false, values: ["01", "02", "03"] }
                    },
                    _TimeBasedValuation: {
                        FixedAssetObjectActionCode: { required: true, values: ["01", "02"] },
                        ValidityStartDate: { required: true, type: "date" },
                        DepreciationKey: { required: true, maxLength: 4 }
                    }
                },

                // Business rules for validation
                businessRules: {
                    minimumChanges: true,
                    accountAssignmentMeaningfulData: true,
                    dateLogicValidation: true,
                    actionCodeConsistency: true
                }
            };

            // Enhanced field mappings for Excel import with Ledger_Area format
            this.fieldMappings = {
                "seq_id": "SequenceNumber",
                "sequencenumber": "SequenceNumber",
                "companycode": "CompanyCode",
                "fixedasset": "FixedAsset",
                "assetnumber": "FixedAsset",
                "asset": "FixedAsset",

                "fixedassetdescription": "FixedAssetDescription",
                "fixedassetdescription_value": "FixedAssetDescription_Value",
                "fixedassetdescription_change": "FixedAssetDescription_Change",

                "assetadditionaldescription": "AssetAdditionalDescription",
                "assetadditionaldescription_value": "AssetAdditionalDescription_Value",
                "assetadditionaldescription_change": "AssetAdditionalDescription_Change",

                "inventory": "Inventory",
                "inventory_value": "Inventory_Value",
                "inventory_change": "Inventory_Change",

                "inventoryiscounted": "InventoryIsCounted",
                "inventoryiscounted_value": "InventoryIsCounted_Value",
                "inventoryiscounted_change": "InventoryIsCounted_Change",

                "assetserialnumber": "AssetSerialNumber",
                "assetserialnumber_value": "AssetSerialNumber_Value",
                "assetserialnumber_change": "AssetSerialNumber_Change",

                "wbselementexternalid": "WBSElementExternalID",
                "wbselementexternalid_value": "WBSElementExternalID_Value",
                "wbselementexternalid_change": "WBSElementExternalID_Change",

                "yy1_wbselement1_faa": "YY1_WBSELEMENT1_FAA",
                "yy1_wbselement1_faa_value": "YY1_WBSELEMENT1_FAA_Value",
                "yy1_wbselement1_faa_change": "YY1_WBSELEMENT1_FAA_Change",

                "costcenter": "CostCenter",
                "costcenter_value": "CostCenter_Value",
                "costcenter_change": "CostCenter_Change",

                "profitcenter": "ProfitCenter",
                "profitcenter_value": "ProfitCenter_Value",
                "profitcenter_change": "ProfitCenter_Change",

                "plant": "Plant",
                "plant_value": "Plant_Value",
                "plant_change": "Plant_Change",

                "room": "Room",
                "room_value": "Room_Value",
                "room_change": "Room_Change",

                "assetlocation": "AssetLocation",
                "assetlocation_value": "AssetLocation_Value",
                "assetlocation_change": "AssetLocation_Change",

                "segment": "Segment",
                "segment_value": "Segment_Value",
                "segment_change": "Segment_Change",

                // v3.0 NEW: remaining AA fields
                "vehiclelicenseplatenumber": "VehicleLicensePlateNumber",
                "vehiclelicenseplatenumber_value": "VehicleLicensePlateNumber_Value",
                "vehiclelicenseplatenumber_change": "VehicleLicensePlateNumber_Change",

                "taxjurisdiction": "TaxJurisdiction",
                "taxjurisdiction_value": "TaxJurisdiction_Value",
                "taxjurisdiction_change": "TaxJurisdiction_Change",

                "businessplace": "BusinessPlace",
                "businessplace_value": "BusinessPlace_Value",
                "businessplace_change": "BusinessPlace_Change",

                "fund": "Fund",
                "fund_value": "Fund_Value",
                "fund_change": "Fund_Change",

                "grantid": "GrantID",
                "grantid_value": "GrantID_Value",
                "grantid_change": "GrantID_Change",

                "functionalarea": "FunctionalArea",
                "functionalarea_value": "FunctionalArea_Value",
                "functionalarea_change": "FunctionalArea_Change",

                "accountassignment_actioncode": "AccountAssignment_ActionCode",
                "accountassignment_validitystartdate": "AccountAssignment_ValidityStartDate",
                "accountassignment_validitystartdate_value": "AccountAssignment_ValidityStartDate_Value",

                "inventorynote": "InventoryNote",
                "inventorynote_value": "InventoryNote_Value",
                "inventorynote_change": "InventoryNote_Change",

                "in_assetblock": "IN_AssetBlock",
                "in_assetblock_value": "IN_AssetBlock_Value",
                "in_assetblock_change": "IN_AssetBlock_Change",

                "in_assetputtousedate": "IN_AssetPutToUseDate",
                "in_assetputtousedate_value": "IN_AssetPutToUseDate_Value",
                "in_assetputtousedate_change": "IN_AssetPutToUseDate_Change",

                "in_assetisprioryear": "IN_AssetIsPriorYear",
                "in_assetisprioryear_value": "IN_AssetIsPriorYear_Value",
                "in_assetisprioryear_change": "IN_AssetIsPriorYear_Change",

                "depreciationkey": "DepreciationKey",
                "plannedusefullifeinyears": "PlannedUsefulLifeInYears",
                "plannedusefullifeinperiods": "PlannedUsefulLifeInPeriods",
                "scrapamountincocodcrcy": "ScrapAmountInCoCodeCrcy",
                "acqnprodnscostcrappercent": "AcqnProdnCostScrapPercent",
                "companycodecurrency": "CompanyCodeCurrency",

                "assetcapitalizationdate": "AssetCapitalizationDate",
                "assetcapitalizationdate_value": "AssetCapitalizationDate_Value",
                "depreciationstartdate": "DepreciationStartDate",
                "depreciationstartdate_value": "DepreciationStartDate_Value",
                "validitystartdate": "ValidityStartDate",
                "validitystartdate_value": "ValidityStartDate_Value",

                "assetcapitalizationdate_0l_01": "AssetCapitalizationDate_0L_01",
                "assetcapitalizationdate_0l_15": "AssetCapitalizationDate_0L_15",
                "assetcapitalizationdate_2l_32": "AssetCapitalizationDate_2L_32",
                "assetcapitalizationdate_3l_34": "AssetCapitalizationDate_3L_34",

                "depreciationstartdate_0l_01": "DepreciationStartDate_0L_01",
                "depreciationstartdate_0l_15": "DepreciationStartDate_0L_15",
                "depreciationstartdate_2l_32": "DepreciationStartDate_2L_32",
                "depreciationstartdate_3l_34": "DepreciationStartDate_3L_34",

                // Ledger action codes (0L only per EDMX)
                "ledger_actioncode_0l_01": "Ledger_ActionCode_0L_01",
                "ledger_actioncode_0l_15": "Ledger_ActionCode_0L_15",

                // Valuation action codes (all areas)
                "valuation_actioncode_0l_01": "Valuation_ActionCode_0L_01",
                "valuation_actioncode_0l_15": "Valuation_ActionCode_0L_15",
                "valuation_actioncode_2l_32": "Valuation_ActionCode_2L_32",
                "valuation_actioncode_3l_34": "Valuation_ActionCode_3L_34",

                // v3.0 NEW: TimeBasedValuation action codes (independent, all areas)
                "timebasedvaluation_actioncode_0l_01": "TimeBasedValuation_ActionCode_0L_01",
                "timebasedvaluation_actioncode_0l_15": "TimeBasedValuation_ActionCode_0L_15",
                "timebasedvaluation_actioncode_2l_32": "TimeBasedValuation_ActionCode_2L_32",
                "timebasedvaluation_actioncode_3l_34": "TimeBasedValuation_ActionCode_3L_34",

                // v3.0 NEW: boolean-only fields (no _Change pair)
                "negativeamountisallowed_0l_01": "NegativeAmountIsAllowed_0L_01",
                "negativeamountisallowed_0l_15": "NegativeAmountIsAllowed_0L_15",
                "negativeamountisallowed_2l_32": "NegativeAmountIsAllowed_2L_32",
                "negativeamountisallowed_3l_34": "NegativeAmountIsAllowed_3L_34",

                "isshutdown_0l_01": "IsShutDown_0L_01",
                "isshutdown_0l_15": "IsShutDown_0L_15",
                "isshutdown_2l_32": "IsShutDown_2L_32",
                "isshutdown_3l_34": "IsShutDown_3L_34",

                // v3.0 NEW: FixedAssetLedger_Type level
                "assetdeactivationdate_0l_01": "AssetDeactivationDate_0L_01",
                "assetdeactivationdate_0l_15": "AssetDeactivationDate_0L_15",
                "assetdeactivationdate_2l_32": "AssetDeactivationDate_2L_32",
                "assetdeactivationdate_3l_34": "AssetDeactivationDate_3L_34",

                // v3.0 NEW: FixedAssetValuation_Type level
                "specialdeprstardate_0l_01": "SpecialDeprStartDate_0L_01",
                "specialdeprstardate_0l_15": "SpecialDeprStartDate_0L_15",
                "specialdeprstardate_2l_32": "SpecialDeprStartDate_2L_32",
                "specialdeprstardate_3l_34": "SpecialDeprStartDate_3L_34",

                "interestcalcstartdate_0l_01": "InterestCalcStartDate_0L_01",
                "interestcalcstartdate_0l_15": "InterestCalcStartDate_0L_15",
                "interestcalcstartdate_2l_32": "InterestCalcStartDate_2L_32",
                "interestcalcstartdate_3l_34": "InterestCalcStartDate_3L_34",

                // Backward compatibility for old area-only format
                "ledger_actioncode_01": "Ledger_ActionCode_0L_01",
                "ledger_actioncode_15": "Ledger_ActionCode_0L_15",

                "valuation_actioncode_01": "Valuation_ActionCode_0L_01",
                "valuation_actioncode_15": "Valuation_ActionCode_0L_15",
                "valuation_actioncode_32": "Valuation_ActionCode_2L_32",
                "valuation_actioncode_34": "Valuation_ActionCode_3L_34"
            };

            // Enhanced changeable fields configuration
            this.changeableFields = {
                // General fields
                "FixedAssetDescription": { section: "_General", allowBlank: false },
                "AssetAdditionalDescription": { section: "_General", allowBlank: true },
                "AssetSerialNumber": { section: "_General", allowBlank: true },

                // Account Assignment fields (YY1_WBSELEMENT1_FAA is actually a root parameter)
                "WBSElementExternalID": { section: "_AccountAssignment", allowBlank: true },
                "CostCenter": { section: "_AccountAssignment", allowBlank: true },
                "ProfitCenter": { section: "_AccountAssignment", allowBlank: true },
                "Segment": { section: "_AccountAssignment", allowBlank: true },
                "Plant": { section: "_AccountAssignment", allowBlank: true },
                "AssetLocation": { section: "_AccountAssignment", allowBlank: true },
                "Room": { section: "_AccountAssignment", allowBlank: true },
                "VehicleLicensePlateNumber": { section: "_AccountAssignment", allowBlank: true },
                "TaxJurisdiction": { section: "_AccountAssignment", allowBlank: true },
                "BusinessPlace": { section: "_AccountAssignment", allowBlank: true },
                "Fund": { section: "_AccountAssignment", allowBlank: true },
                "GrantID": { section: "_AccountAssignment", allowBlank: true },
                "FunctionalArea": { section: "_AccountAssignment", allowBlank: true },

                "YY1_WBSELEMENT1_FAA": { section: "_Root", allowBlank: true },

                // Inventory fields
                "InventoryNote": { section: "_Inventory", allowBlank: true },
                "Inventory": { section: "_Inventory", allowBlank: true },
                "InventoryIsCounted": { section: "_Inventory", allowBlank: true },

                // India-specific fields
                "IN_AssetBlock": { section: "_GlobMasterData", allowBlank: true },
                "IN_AssetPutToUseDate": { section: "_GlobMasterData", allowBlank: true },
                "IN_AssetIsPriorYear": { section: "_GlobMasterData", allowBlank: true },

                // Depreciation fields
                "AssetCapitalizationDate": { section: "_Ledger", allowBlank: true },
                "DepreciationStartDate": { section: "_Ledger", allowBlank: true },
                "PlannedUsefulLifeInYears": { section: "_Ledger", allowBlank: false },
                "PlannedUsefulLifeInPeriods": { section: "_Ledger", allowBlank: true },
                "ScrapAmountInCoCodeCrcy": { section: "_Ledger", allowBlank: true },
                "AcqnProdnCostScrapPercent": { section: "_Ledger", allowBlank: true }
            };

            // Enhanced field definitions for display with Ledger_Area format

            this.FIELD_DEFINITIONS = [
                // General Information
                { key: 'FixedAssetDescription', section: 'General', displayName: 'Asset Description' },
                { key: 'AssetAdditionalDescription', section: 'General', displayName: 'Additional Description' },
                { key: 'AssetSerialNumber', section: 'General', displayName: 'Serial Number' },

                // India-specific
                { key: 'IN_AssetBlock', section: 'General', displayName: 'Asset Block' },
                { key: 'IN_AssetPutToUseDate', section: 'General', displayName: 'Put to Use Date' },
                { key: 'IN_AssetIsPriorYear', section: 'General', displayName: 'Is Prior Year' },
                { key: 'IN_AdditionalAssetBlock', section: 'General', displayName: 'Additional Asset Block' },
                { key: 'IN_AssetIsResearchAndDev', section: 'General', displayName: 'Is Research & Development' },

                // Account Assignment
                { key: 'AccountAssignment_ActionCode', section: 'AccountAssignment', displayName: 'Account Assignment Action' },
                { key: 'WBSElementExternalID', section: 'AccountAssignment', displayName: 'WBS Element' },
                { key: 'YY1_WBSELEMENT1_FAA', section: 'AccountAssignment', displayName: 'YY1 WBS Element' },
                { key: 'CostCenter', section: 'AccountAssignment', displayName: 'Cost Center' },
                { key: 'ProfitCenter', section: 'AccountAssignment', displayName: 'Profit Center' },
                { key: 'Plant', section: 'AccountAssignment', displayName: 'Plant' },
                { key: 'AssetLocation', section: 'AccountAssignment', displayName: 'Asset Location' },
                { key: 'Room', section: 'AccountAssignment', displayName: 'Room' },
                { key: 'Segment', section: 'AccountAssignment', displayName: 'Segment' },
                // v3.0 NEW: remaining AA fields
                { key: 'VehicleLicensePlateNumber', section: 'AccountAssignment', displayName: 'Vehicle License Plate' },
                { key: 'TaxJurisdiction', section: 'AccountAssignment', displayName: 'Tax Jurisdiction' },
                { key: 'BusinessPlace', section: 'AccountAssignment', displayName: 'Business Place' },
                { key: 'Fund', section: 'AccountAssignment', displayName: 'Fund' },
                { key: 'GrantID', section: 'AccountAssignment', displayName: 'Grant ID' },
                { key: 'FunctionalArea', section: 'AccountAssignment', displayName: 'Functional Area' },

                // Inventory
                { key: 'Inventory', section: 'Inventory', displayName: 'Inventory Number' },
                { key: 'InventoryNote', section: 'Inventory', displayName: 'Inventory Note' },
                { key: 'InventoryIsCounted', section: 'Inventory', displayName: 'Is Inventory Counted' },

                // Depreciation Areas — dynamic per area, Ledger_ActionCode only for 0L
                ...['0L_01', '0L_15', '2L_32', '3L_34'].flatMap(ledgerArea => {
                    const ledger = ledgerArea.split('_')[0];
                    const hasLedgerActionCode = ['0L'].includes(ledger); // per EDMX: 0L only
                    return [
                        ...(hasLedgerActionCode ? [
                            { key: `Ledger_ActionCode_${ledgerArea}`, section: 'Depreciation', displayName: `Ledger Action (${ledgerArea})` }
                        ] : []),
                        { key: `Valuation_ActionCode_${ledgerArea}`, section: 'Depreciation', displayName: `Valuation Action (${ledgerArea})` },
                        { key: `TimeBasedValuation_ActionCode_${ledgerArea}`, section: 'Depreciation', displayName: `TBV Action (${ledgerArea})` },           // v3.0 NEW
                        { key: `NegativeAmountIsAllowed_${ledgerArea}`, section: 'Depreciation', displayName: `Negative Amount Allowed (${ledgerArea})` },    // v3.0 NEW
                        { key: `AssetCapitalizationDate_${ledgerArea}`, section: 'Depreciation', displayName: `Asset Capitalization Date (${ledgerArea})` },
                        { key: `AssetDeactivationDate_${ledgerArea}`, section: 'Depreciation', displayName: `Asset Deactivation Date (${ledgerArea})` },       // v3.0 NEW
                        { key: `DepreciationStartDate_${ledgerArea}`, section: 'Depreciation', displayName: `Depreciation Start Date (${ledgerArea})` },
                        { key: `SpecialDeprStartDate_${ledgerArea}`, section: 'Depreciation', displayName: `Special Depr Start Date (${ledgerArea})` },        // v3.0 NEW
                        { key: `InterestCalcStartDate_${ledgerArea}`, section: 'Depreciation', displayName: `Interest Calc Start Date (${ledgerArea})` },      // v3.0 NEW
                        { key: `DepreciationKey_${ledgerArea}`, section: 'Depreciation', displayName: `Depreciation Key (${ledgerArea})` },
                        { key: `PlannedUsefulLifeInYears_${ledgerArea}`, section: 'Depreciation', displayName: `Useful Life Years (${ledgerArea})` },
                        { key: `PlannedUsefulLifeInPeriods_${ledgerArea}`, section: 'Depreciation', displayName: `Useful Life Periods (${ledgerArea})` },
                        { key: `ScrapAmountInCoCodeCrcy_${ledgerArea}`, section: 'Depreciation', displayName: `Scrap Amount (${ledgerArea})` },
                        { key: `AcqnProdnCostScrapPercent_${ledgerArea}`, section: 'Depreciation', displayName: `Scrap Percentage (${ledgerArea})` },
                        { key: `VariableDeprPercent_${ledgerArea}`, section: 'Depreciation', displayName: `Variable Depr % (${ledgerArea})` },                 // v3.0 NEW
                        { key: `DeprCalcBaseValuePercent_${ledgerArea}`, section: 'Depreciation', displayName: `Depr Calc Base % (${ledgerArea})` },           // v3.0 NEW
                        { key: `ShiftOperationFactor_${ledgerArea}`, section: 'Depreciation', displayName: `Shift Factor (${ledgerArea})` },                   // v3.0 NEW
                        { key: `IsShutDown_${ledgerArea}`, section: 'Depreciation', displayName: `Is Shut Down (${ledgerArea})` },                            // v3.0 NEW
                        { key: `AssetRevaluationIndex_${ledgerArea}`, section: 'Depreciation', displayName: `Revaluation Index (${ledgerArea})` }             // v3.0 NEW
                    ];
                })
            ];
            // Display name mappings
            this.DISPLAY_NAMES = this.FIELD_DEFINITIONS.reduce((acc, field) => {
                acc[field.key] = field.displayName;
                const baseField = field.key.replace(/_\w+_\w+$/, '');
                if (!acc[baseField]) {
                    acc[baseField] = field.displayName.replace(/ \(\w+_\w+\)$/, '');
                }
                return acc;
            }, {});
        },

        // =================
        // HELPER METHODS FOR DATE EXTRACTION
        // =================

        /**
         * Extract date value from Excel data if present
         * @private
         */
        _extractDateFromExcel: function (data, fieldName) {
            // Try direct field name first
            if (data[fieldName] && this._isValidDate(data[fieldName])) {
                return this._formatDate(data[fieldName]);
            }

            // Try with _Value suffix
            const valueKey = fieldName + '_Value';
            if (data[valueKey] && this._isValidDate(data[valueKey])) {
                return this._formatDate(data[valueKey]);
            }

            // Return null if no valid date found in Excel
            return null;
        },

        /**
         * Check if a value is a valid date
         * @private
         */
        _isValidDate: function (value) {
            if (!value || value === "" || typeof value === 'undefined') {
                return false;
            }

            const date = new Date(value);
            return !isNaN(date.getTime());
        },

        /**
         * Format date to YYYY-MM-DD
         * @private
         */
        _formatDate: function (dateValue) {
            if (!dateValue) return null;

            try {
                const date = new Date(dateValue);
                if (isNaN(date.getTime())) return null;

                return date.toISOString().split('T')[0];
            } catch (error) {
                console.warn("Date formatting failed:", dateValue, error);
                return null;
            }
        },

        // =================
        // CHANGE TRANSFORMATION WITH LEDGER_AREA FORMAT
        // =================

        convertFlatToChangeStructure: function (flatRow) {
            console.log("[DataTransformer] Converting flat row to change structure with enhanced validation");

            const mapped = this._mapBasicFields(flatRow);
            const changeStructure = this._buildChangeStructure(mapped);

            try {
                // Run comprehensive validation
                changeStructure.ValidationErrors = this.validateAssetChangeData(changeStructure);

                // Determine status based on validation results
                const errorCount = changeStructure.ValidationErrors.length;
                const criticalErrors = changeStructure.ValidationErrors.filter(error =>
                    error.severity === 'Error' || error.validationType === 'MANDATORY_FIELD'
                ).length;

                if (criticalErrors > 0) {
                    changeStructure.Status = 'Invalid';
                } else if (errorCount > 0) {
                    changeStructure.Status = 'Warning'; // Has warnings but processable
                } else {
                    changeStructure.Status = 'Valid';
                }

                // Add validation summary
                changeStructure.ValidationSummary = {
                    totalErrors: errorCount,
                    criticalErrors: criticalErrors,
                    warnings: errorCount - criticalErrors,
                    isProcessable: criticalErrors === 0
                };

            } catch (error) {
                changeStructure.ValidationErrors = [{
                    field: 'validation',
                    message: `Validation error: ${error.message}`,
                    severity: 'Error',
                    validationType: 'SYSTEM_ERROR',
                    sequenceId: changeStructure.SequenceNumber || 'N/A',
                    asset: changeStructure.FixedAsset || 'Unknown'
                }];
                changeStructure.Status = 'Invalid';
            }

            return changeStructure;
        },

        _buildChangeStructure: function (data) {
            const result = {
                SequenceNumber: data.SequenceNumber,
                CompanyCode: data.CompanyCode,
                FixedAsset: data.FixedAsset,
                changes: {},
                actionCodes: {
                    accountAssignment: null,
                    ledgers: {},
                    valuations: {},
                    timeBasedValuations: {}  // v3.0 NEW
                },
                _originalExcelData: data
            };

            this._processActionCodesWithLedgerArea(data, result);
            this._processValueChangePairs(data, result);
            this._processLedgerAreaBasedChanges(data, result);
            this._processBooleanOnlyFields(data, result); // v3.0 NEW

            return result;
        },

        _processActionCodesWithLedgerArea: function (data, result) {
            // Account Assignment action code
            if (data.AccountAssignment_ActionCode) {
                const actionCode = String(data.AccountAssignment_ActionCode).trim();
                if (this.config.validActionsByLevel.accountAssignment.includes(actionCode)) {
                    result.actionCodes.accountAssignment = actionCode;
                    console.log(`[DataTransformer] Account Assignment action code: ${actionCode}`);
                }
            }

            Object.keys(this.config.ledgerAreaToLedgerMap).forEach(ledgerArea => {
                const ledger = ledgerArea.split('_')[0];

                // Ledger action code — only valid for 0L per EDMX
                if (this.config.ledgerActionCodeLedgers.includes(ledger)) {
                    const ledgerActionKey = `Ledger_ActionCode_${ledgerArea}`;
                    if (data[ledgerActionKey]) {
                        const actionCode = String(data[ledgerActionKey]).trim();
                        if (this.config.validActionsByLevel.ledger.includes(actionCode)) {
                            result.actionCodes.ledgers[ledgerArea] = actionCode;
                            console.log(`[DataTransformer] Ledger action code for ${ledgerArea}: ${actionCode}`);
                        }
                    }
                }

                // Valuation action code (all areas)
                const valuationActionKey = `Valuation_ActionCode_${ledgerArea}`;
                if (data[valuationActionKey]) {
                    const actionCode = String(data[valuationActionKey]).trim();
                    if (this.config.validActionsByLevel.valuation.includes(actionCode)) {
                        result.actionCodes.valuations[ledgerArea] = actionCode;
                        console.log(`[DataTransformer] Valuation action code for ${ledgerArea}: ${actionCode}`);
                    }
                }

                // v3.0 NEW: TimeBasedValuation action code (independent from Valuation)
                const timeBasedActionKey = `TimeBasedValuation_ActionCode_${ledgerArea}`;
                if (data[timeBasedActionKey]) {
                    const actionCode = String(data[timeBasedActionKey]).trim();
                    if (this.config.validActionsByLevel.timeBasedValuation.includes(actionCode)) {
                        result.actionCodes.timeBasedValuations[ledgerArea] = actionCode;
                        console.log(`[DataTransformer] TimeBasedValuation action code for ${ledgerArea}: ${actionCode}`);
                    }
                }
            });
        },

        _processValueChangePairs: function (data, result) {
            Object.keys(data).forEach(key => {
                if (key.endsWith('_Change') && !key.includes('ActionCode')) {
                    const baseFieldName = key.replace('_Change', '');
                    const valueKey = baseFieldName + '_Value';
                    const changeFlag = this._parseChangeFlag(data[key]);

                    if (changeFlag && data.hasOwnProperty(valueKey)) {
                        const rawValue = data[valueKey];
                        const processedValue = this._processChangeValue(rawValue);

                        result.changes[baseFieldName] = {
                            newValue: processedValue.value,
                            shouldChange: true,
                            isBlankValue: processedValue.isBlank,
                            originalValue: rawValue
                        };
                    }
                }
            });
        },

        _processLedgerAreaBasedChanges: function (data, result) {
            Object.keys(data).forEach(key => {
                const patternMatch = key.match(/^(.+?)_(\w+_\w+)_Change$/i);
                if (patternMatch && !key.includes('ActionCode')) {
                    const [, fieldName, ledgerAreaCode] = patternMatch;

                    if (!this.config.ledgerAreaToLedgerMap[ledgerAreaCode]) {
                        return;
                    }

                    const valueKey = `${fieldName}_${ledgerAreaCode}_Value`;
                    const changeFlag = this._parseChangeFlag(data[key]);

                    if (changeFlag && data.hasOwnProperty(valueKey)) {
                        const rawValue = data[valueKey];
                        const processedValue = this._processChangeValue(rawValue);
                        const fullFieldName = `${fieldName}_${ledgerAreaCode}`;

                        const ledgerActionCode = result.actionCodes.ledgers[ledgerAreaCode] || this.config.defaultActionCode;
                        const valuationActionCode = result.actionCodes.valuations[ledgerAreaCode] || this.config.defaultActionCode;
                        const timeBasedValuationActionCode = result.actionCodes.timeBasedValuations[ledgerAreaCode] || null; // v3.0 NEW

                        const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerAreaCode];
                        const areaCode = ledgerAreaMapping ? ledgerAreaMapping.area : ledgerAreaCode.split('_')[1];
                        const ledgerCode = ledgerAreaMapping ? ledgerAreaMapping.ledger : ledgerAreaCode.split('_')[0];

                        result.changes[fullFieldName] = {
                            newValue: processedValue.value,
                            shouldChange: true,
                            isBlankValue: processedValue.isBlank,
                            originalValue: rawValue,
                            areaCode: areaCode,
                            ledgerCode: ledgerCode,
                            ledgerAreaCode: ledgerAreaCode,
                            baseField: fieldName,
                            ledgerActionCode: ledgerActionCode,
                            valuationActionCode: valuationActionCode,
                            timeBasedValuationActionCode: timeBasedValuationActionCode // v3.0 NEW
                        };
                    }
                }
            });
        },
        /**
         * v3.0 NEW: Process boolean-only fields that have no _Change pair.
         * NegativeAmountIsAllowed and IsShutDown are set directly — no _Change column.
         * @private
         */
        _processBooleanOnlyFields: function (data, result) {
            const booleanOnlyFields = ["NegativeAmountIsAllowed", "IsShutDown"];

            booleanOnlyFields.forEach(fieldBase => {
                Object.keys(this.config.ledgerAreaToLedgerMap).forEach(ledgerArea => {
                    const colKey = `${fieldBase}_${ledgerArea}`;
                    const rawVal = data[colKey];

                    // Only process if a non-empty value was provided
                    if (rawVal === null || rawVal === undefined || rawVal === "") return;

                    const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerArea];
                    const areaCode = ledgerAreaMapping ? ledgerAreaMapping.area : ledgerArea.split('_')[1];
                    const ledgerCode = ledgerAreaMapping ? ledgerAreaMapping.ledger : ledgerArea.split('_')[0];
                    const valuationActionCode = result.actionCodes.valuations[ledgerArea] || this.config.defaultActionCode;
                    const timeBasedValuationActionCode = result.actionCodes.timeBasedValuations[ledgerArea] || null;

                    result.changes[colKey] = {
                        newValue: this._parseBoolean(rawVal),
                        shouldChange: true,
                        isBlankValue: false,
                        isBooleanOnly: true, // v3.0 marker: no _Change pair
                        originalValue: rawVal,
                        areaCode: areaCode,
                        ledgerCode: ledgerCode,
                        ledgerAreaCode: ledgerArea,
                        baseField: fieldBase,
                        ledgerActionCode: result.actionCodes.ledgers[ledgerArea] || this.config.defaultActionCode,
                        valuationActionCode: valuationActionCode,
                        timeBasedValuationActionCode: timeBasedValuationActionCode
                    };

                    console.log(`[DataTransformer] Boolean-only field ${colKey} = ${this._parseBoolean(rawVal)}`);
                });
            });
        },
        validateAssetChangeData: function (changeData) {
            const errors = [];
            const errorKeys = new Set();
            console.log("[DataTransformer] Starting enhanced validation with mandatory field checks");

            const addUniqueError = (error) => {
                const errorKey = `${error.field}_${error.validationType}_${error.message}`;
                if (!errorKeys.has(errorKey)) {
                    errorKeys.add(errorKey);
                    errors.push(error);
                }
            };

            // Phase 1: Basic required fields validation
            this._validateBasicRequiredFields(changeData, errors, addUniqueError);

            // Phase 2: Section-specific mandatory field validation
            this._validateSectionMandatoryFields(changeData, errors, addUniqueError);

            // Phase 3: Field format and length validation
            this._validateFieldFormatsAndLengths(changeData, errors, addUniqueError);

            // Phase 4: Business logic validation
            this._validateBusinessLogic(changeData, errors, addUniqueError);

            // Phase 5: Action code consistency validation
            this._validateActionCodeConsistency(changeData, errors, addUniqueError);

            // ✅ NEW: Phase 6: ValidityStartDate requirements validation
            this._validateValidityStartDateRequirements(changeData, errors, addUniqueError);

            console.log(`[DataTransformer] Validation completed with ${errors.length} unique errors`);
            return errors;
        },

        // UPDATED: All validation methods to use addUniqueError function
        _validateBasicRequiredFields: function (changeData, errors, addUniqueError = null) {
            const requiredFields = this.validationConfig.mandatoryFields._Root;
            const addError = addUniqueError || ((error) => errors.push(error));

            Object.keys(requiredFields).forEach(fieldName => {
                const fieldConfig = requiredFields[fieldName];
                const value = this._getValue(changeData, fieldName);

                if (fieldConfig.required && (!value || String(value).trim() === "")) {
                    addError({
                        field: fieldName,
                        message: `${fieldName} is mandatory and cannot be empty`,
                        severity: 'Error',
                        sequenceId: changeData.SequenceNumber || 'N/A',
                        asset: changeData.FixedAsset || 'Unknown',
                        validationType: 'MANDATORY_FIELD'
                    });
                }

                // Validate field length
                if (value && fieldConfig.maxLength && String(value).length > fieldConfig.maxLength) {
                    addError({
                        field: fieldName,
                        message: `${fieldName} exceeds maximum length of ${fieldConfig.maxLength} characters`,
                        severity: 'Error',
                        sequenceId: changeData.SequenceNumber || 'N/A',
                        asset: changeData.FixedAsset || 'Unknown',
                        validationType: 'FIELD_LENGTH'
                    });
                }
            });
        },

        // UPDATED: Section-specific validation with duplicate prevention
        _validateSectionMandatoryFields: function (changeData, errors, addUniqueError = null) {
            const changes = changeData.changes || {};
            const actionCodes = changeData.actionCodes || {};
            const addError = addUniqueError || ((error) => errors.push(error));

            // Check if AccountAssignment section has changes
            const hasAccountAssignmentChanges = Object.keys(changes).some(fieldName =>
                this._getFieldSection(fieldName) === '_AccountAssignment'
            );

            if (hasAccountAssignmentChanges || actionCodes.accountAssignment) {
                this._validateAccountAssignmentMandatoryFields(changeData, errors, addError);
            }

            // Check Ledger sections
            Object.keys(actionCodes.ledgers || {}).forEach(ledgerArea => {
                this._validateLedgerMandatoryFields(changeData, ledgerArea, errors, addError);
            });

            // Check Valuation sections
            Object.keys(actionCodes.valuations || {}).forEach(ledgerArea => {
                this._validateValuationMandatoryFields(changeData, ledgerArea, errors, addError);
            });
        },

        // UPDATED: AccountAssignment validation with duplicate prevention
        _validateAccountAssignmentMandatoryFields: function (changeData, errors, addError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const addErrorFn = addError || ((error) => errors.push(error));

            // FixedAssetObjectActionCode is mandatory for AccountAssignment
            const actionCode = changeData.actionCodes?.accountAssignment;
            if (!actionCode) {
                addErrorFn({
                    field: 'AccountAssignment_ActionCode',
                    message: 'Account Assignment Action Code is mandatory when changing account assignment fields',
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'MANDATORY_FIELD'
                });
            } else if (!this.validationConfig.mandatoryFields._AccountAssignment.FixedAssetObjectActionCode.values.includes(actionCode)) {
                addErrorFn({
                    field: 'AccountAssignment_ActionCode',
                    message: `Invalid Account Assignment Action Code '${actionCode}'. Valid values: 01 (Create), 02 (Change)`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'INVALID_VALUE'
                });
            }

            // Check if AccountAssignment has meaningful data
            const changes = changeData.changes || {};
            const accountAssignmentChanges = Object.keys(changes).filter(fieldName =>
                this._getFieldSection(fieldName) === '_AccountAssignment'
            );

            const meaningfulChanges = accountAssignmentChanges.filter(fieldName => {
                const change = changes[fieldName];
                return change.shouldChange && !change.isBlankValue && change.newValue && String(change.newValue).trim() !== '';
            });

            if (accountAssignmentChanges.length > 0 && meaningfulChanges.length === 0) {
                addErrorFn({
                    field: 'AccountAssignment',
                    message: 'Account Assignment changes require at least one non-blank field value',
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'BUSINESS_LOGIC'
                });
            }
        },

        // UPDATED: Other validation methods with addError parameter
        _validateLedgerMandatoryFields: function (changeData, ledgerArea, errors, addError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const addErrorFn = addError || ((error) => errors.push(error));

            // Extract ledger code from ledgerArea
            const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerArea];
            if (!ledgerAreaMapping) {
                addErrorFn({
                    field: `Ledger_${ledgerArea}`,
                    message: `Invalid Ledger Area code: ${ledgerArea}`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'INVALID_VALUE'
                });
                return;
            }

            const actionCode = changeData.actionCodes.ledgers[ledgerArea];

            // Validate action code if present
            if (actionCode && !this.validationConfig.mandatoryFields._Ledger.FixedAssetObjectActionCode.values.includes(actionCode)) {
                addErrorFn({
                    field: `Ledger_ActionCode_${ledgerArea}`,
                    message: `Invalid Ledger Action Code '${actionCode}' for ${ledgerArea}. Valid values: 01 (Create), 02 (Change), 03 (Delete)`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'INVALID_VALUE'
                });
            }
        },

        // UPDATED: Valuation validation with addError parameter
        _validateValuationMandatoryFields: function (changeData, ledgerArea, errors, addError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const addErrorFn = addError || ((error) => errors.push(error));

            const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerArea];
            if (!ledgerAreaMapping) return;

            const actionCode = changeData.actionCodes.valuations[ledgerArea];

            // Validate action code if present
            if (actionCode && !this.validationConfig.mandatoryFields._Valuation.FixedAssetObjectActionCode.values.includes(actionCode)) {
                addErrorFn({
                    field: `Valuation_ActionCode_${ledgerArea}`,
                    message: `Invalid Valuation Action Code '${actionCode}' for ${ledgerArea}. Valid values: 01 (Create), 02 (Change), 03 (Delete)`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'INVALID_VALUE'
                });
            }

            // For DELETE action, no time-based valuation fields should be present
            if (actionCode === "03") {
                const hasTimeBasedFields = Object.keys(changeData.changes || {}).some(fieldName =>
                    fieldName.includes(`_${ledgerArea}`) &&
                    !fieldName.includes(`Valuation_ActionCode_${ledgerArea}`) &&
                    !fieldName.includes(`Ledger_ActionCode_${ledgerArea}`)
                );

                if (hasTimeBasedFields) {
                    addErrorFn({
                        field: `Valuation_ActionCode_${ledgerArea}`,
                        message: `When deleting valuation (action code 03), no time-based valuation fields should be provided for ${ledgerArea}`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset,
                        validationType: 'BUSINESS_LOGIC'
                    });
                }
            }

            // For CREATE/CHANGE actions, check for required depreciation fields
            if (actionCode === "01" || actionCode === "02") {
                const requiredDepreciationFields = ['DepreciationKey'];

                requiredDepreciationFields.forEach(fieldName => {
                    const fullFieldName = `${fieldName}_${ledgerArea}`;
                    const fieldChange = changeData.changes[fullFieldName];

                    if (!fieldChange || !fieldChange.shouldChange || fieldChange.isBlankValue) {
                        addErrorFn({
                            field: fullFieldName,
                            message: `${fieldName} is required for valuation ${actionCode === "01" ? "creation" : "change"} in ${ledgerArea}`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'MANDATORY_FIELD'
                        });
                    }
                });
            }
        },

        // UPDATED: Field format validation with addError parameter
        _validateFieldFormatsAndLengths: function (changeData, errors, addUniqueError = null) {
            const changes = changeData.changes || {};
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const addError = addUniqueError || ((error) => errors.push(error));

            // Field length limits from SAP API documentation
            const fieldLimits = {
                FixedAssetDescription: { maxLength: 50 },
                AssetAdditionalDescription: { maxLength: 50 },
                AssetSerialNumber: { maxLength: 18 },
                WBSElementExternalID: { maxLength: 24 },
                CostCenter: { maxLength: 10 },
                ProfitCenter: { maxLength: 10 },
                Plant: { maxLength: 4 },
                Room: { maxLength: 8 },
                AssetLocation: { maxLength: 10 },
                VehicleLicensePlateNumber: { maxLength: 15 },
                TaxJurisdiction: { maxLength: 15 },
                BusinessPlace: { maxLength: 4 },
                Fund: { maxLength: 10 },
                GrantID: { maxLength: 20 },
                FunctionalArea: { maxLength: 16 },
                Inventory: { maxLength: 25 },
                InventoryNote: { maxLength: 15 },
                DepreciationKey: { maxLength: 4, exactLength: 4 },
                PlannedUsefulLifeInYears: { maxLength: 3 },
                PlannedUsefulLifeInPeriods: { maxLength: 3 }
            };

            Object.entries(changes).forEach(([fieldName, changeInfo]) => {
                if (!changeInfo.shouldChange || changeInfo.isBlankValue) return;

                const value = changeInfo.newValue;
                const baseFieldName = changeInfo.baseField || fieldName.replace(/_\w+_\w+$/, '');
                const fieldLimits_config = fieldLimits[baseFieldName];

                if (fieldLimits_config && value) {
                    // Check length limits
                    if (fieldLimits_config.maxLength && String(value).length > fieldLimits_config.maxLength) {
                        addError({
                            field: fieldName,
                            message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} exceeds maximum length of ${fieldLimits_config.maxLength} characters (current: ${String(value).length})`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'FIELD_LENGTH'
                        });
                    }

                    // Check exact length requirements
                    if (fieldLimits_config.exactLength && String(value).length !== fieldLimits_config.exactLength) {
                        addError({
                            field: fieldName,
                            message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be exactly ${fieldLimits_config.exactLength} characters`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'FIELD_FORMAT'
                        });
                    }
                }
            });
        },

        // UPDATED: Business logic validation with addError parameter
        _validateBusinessLogic: function (changeData, errors, addUniqueError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const changes = changeData.changes || {};
            const addError = addUniqueError || ((error) => errors.push(error));

            // Rule 1: If no changes specified, this is invalid
            const hasAnyChanges = Object.values(changes).some(change => change.shouldChange);
            if (!hasAnyChanges) {
                addError({
                    field: 'changes',
                    message: 'At least one field must be marked for change',
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset,
                    validationType: 'BUSINESS_LOGIC'
                });
            }

            // Rule 2: Validate date logic
            this._validateDateLogic(changeData, errors, addError);
        },

        // UPDATED: Date logic validation with addError parameter
        _validateDateLogic: function (changeData, errors, addError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const addErrorFn = addError || ((error) => errors.push(error));

            // Check AssetCapitalizationDate vs DepreciationStartDate logic
            Object.keys(this.config.ledgerAreaToLedgerMap).forEach(ledgerArea => {
                const capitalizationDate = this._extractDateFromExcel(
                    changeData._originalExcelData,
                    `AssetCapitalizationDate_${ledgerArea}`
                );
                const depreciationStartDate = this._extractDateFromExcel(
                    changeData._originalExcelData,
                    `DepreciationStartDate_${ledgerArea}`
                );

                if (capitalizationDate && depreciationStartDate) {
                    const capDate = new Date(capitalizationDate);
                    const deprDate = new Date(depreciationStartDate);

                    if (deprDate < capDate) {
                        addErrorFn({
                            field: `DepreciationStartDate_${ledgerArea}`,
                            message: `Depreciation Start Date cannot be earlier than Asset Capitalization Date for ${ledgerArea}`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'BUSINESS_LOGIC'
                        });
                    }
                }
            });
        },

        // UPDATED: Action code consistency validation with addError parameter
        _validateActionCodeConsistency: function (changeData, errors, addUniqueError = null) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';
            const actionCodes = changeData.actionCodes || {};
            const addError = addUniqueError || ((error) => errors.push(error));

            // Check consistency between ledger and valuation action codes
            Object.keys(this.config.ledgerAreaToLedgerMap).forEach(ledgerArea => {
                const ledgerAction = actionCodes.ledgers?.[ledgerArea];
                const valuationAction = actionCodes.valuations?.[ledgerArea];

                // If deleting ledger, valuation should also be deleted or not specified
                if (ledgerAction === "03" && valuationAction && valuationAction !== "03") {
                    addError({
                        field: `ActionCode_${ledgerArea}`,
                        message: `When deleting ledger ${ledgerArea}, valuation should also be deleted or not specified`,
                        severity: 'Warning',
                        sequenceId: sequenceId,
                        asset: asset,
                        validationType: 'ACTION_CODE_CONSISTENCY'
                    });
                }

                // If creating ledger, valuation should be created or changed
                if (ledgerAction === "01" && valuationAction === "03") {
                    addError({
                        field: `ActionCode_${ledgerArea}`,
                        message: `When creating ledger ${ledgerArea}, valuation cannot be deleted`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset,
                        validationType: 'ACTION_CODE_CONSISTENCY'
                    });
                }
            });
        },

        _validateActionCodesWithLedgerArea: function (changeData, errors) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';

            // Validate Account Assignment action code
            if (changeData.actionCodes.accountAssignment) {
                const code = changeData.actionCodes.accountAssignment;
                if (!this.config.validActionsByLevel.accountAssignment.includes(code)) {
                    errors.push({
                        field: 'AccountAssignment_ActionCode',
                        message: `Invalid Account Assignment action code '${code}'. Valid codes are: 01 (Create), 02 (Change)`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                }
            }

            // Validate Ledger action codes with Ledger_Area format
            Object.entries(changeData.actionCodes.ledgers).forEach(([ledgerArea, code]) => {
                if (!this.config.validActionsByLevel.ledger.includes(code)) {
                    errors.push({
                        field: `Ledger_ActionCode_${ledgerArea}`,
                        message: `Invalid Ledger action code '${code}' for ${ledgerArea}. Valid codes are: 01 (Create), 02 (Change), 03 (Delete)`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                }
            });

            // Validate Valuation action codes with Ledger_Area format
            Object.entries(changeData.actionCodes.valuations).forEach(([ledgerArea, code]) => {
                if (!this.config.validActionsByLevel.valuation.includes(code)) {
                    errors.push({
                        field: `Valuation_ActionCode_${ledgerArea}`,
                        message: `Invalid Valuation action code '${code}' for ${ledgerArea}. Valid codes are: 01 (Create), 02 (Change), 03 (Delete)`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                }

                // Additional validation: if deleting valuation, no time-based valuation fields should be set
                if (code === "03") {
                    const hasDepreciationFields = Object.keys(changeData.changes).some(field =>
                        field.includes(`_${ledgerArea}`) && !field.includes(`Valuation_ActionCode_${ledgerArea}`)
                    );
                    if (hasDepreciationFields) {
                        errors.push({
                            field: `Valuation_ActionCode_${ledgerArea}`,
                            message: `When deleting valuation (action code 03), no depreciation fields should be provided for ${ledgerArea}`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset
                        });
                    }
                }
            });
        },

        _validateFieldChangeWithLedgerArea: function (fieldName, changeInfo, changeData) {
            const errors = [];
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';

            const baseFieldName = changeInfo.baseField || fieldName;

            // DEBUG: Log the field lookup process
            console.log("[DEBUG] Field validation:", {
                fieldName: fieldName,
                baseFieldName: baseFieldName,
                changeInfoBaseField: changeInfo.baseField,
                availableFields: Object.keys(this.changeableFields),
                lookupResult: this.changeableFields[baseFieldName]
            });

            const fieldConfig = this.changeableFields[baseFieldName];

            if (!fieldConfig) {
                console.error(`[ERROR] Field '${fieldName}' (baseField: '${baseFieldName}') not found in changeableFields`);
                errors.push({
                    field: fieldName,
                    message: `Field '${fieldName}' is not changeable`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset
                });
                return errors;
            }

            // Get the appropriate action code based on field section
            let actionCode = this.config.defaultActionCode;
            if (fieldConfig.section === '_AccountAssignment' && changeData.actionCodes.accountAssignment) {
                actionCode = changeData.actionCodes.accountAssignment;
            } else if (fieldConfig.section === '_Ledger' && changeInfo.ledgerAreaCode) {
                actionCode = changeInfo.valuationActionCode || this.config.defaultActionCode;
            }

            // For DELETE action at valuation level, most validations don't apply
            if (actionCode === this.config.actionCodes.delete && fieldConfig.section === '_Ledger') {
                return errors; // No validation needed for delete
            }

            // For CREATE action, blank values are not allowed for required fields
            if (actionCode === this.config.actionCodes.create && changeInfo.isBlankValue && !fieldConfig.allowBlank) {
                errors.push({
                    field: fieldName,
                    message: `Field '${fieldName}' cannot be blank when creating new data`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset
                });
            }

            // For CHANGE action, apply normal validation
            if (actionCode === this.config.actionCodes.change) {
                if (changeInfo.isBlankValue && !fieldConfig.allowBlank) {
                    errors.push({
                        field: fieldName,
                        message: `Field '${fieldName}' cannot be set to blank`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                }
            }

            if (!changeInfo.isBlankValue && changeInfo.newValue) {
                this._validateFieldValue(fieldName, changeInfo.newValue, errors, sequenceId, asset);
            }

            return errors;
        },

        // =================
        // ENHANCED ODATA TRANSFORMATION WITH LEDGER_AREA FORMAT - FIXED VERSION
        // =================

        transformToODataFormat: function (changeData, existingSapData = null) {
            console.log("[DataTransformer] Transforming change data to OData format with metadata validation");

            if (!changeData || !changeData.changes) {
                throw new Error("Invalid change data provided");
            }

            const payload = {};

            // Process changes using enhanced helper methods
            Object.entries(changeData.changes).forEach(([fieldName, changeInfo]) => {
                if (!changeInfo.shouldChange) return;

                // Use existing _getFieldSection method
                const section = this._getFieldSection(fieldName);
                const value = changeInfo.isBlankValue ? "" : changeInfo.newValue;

                // Build payload structure based on existing field definitions
                this._buildODataSectionWithLedgerArea(payload, section, fieldName, value, changeInfo, changeData);
            });

            // CRITICAL: Ensure proper structure for empty sections based on metadata
            if (payload._Ledger && payload._Ledger.length > 0) {
                payload._Ledger.forEach(ledger => {
                    if (ledger._Valuation && ledger._Valuation.length > 0) {
                        ledger._Valuation.forEach(valuation => {
                            if (!valuation._TimeBasedValuation || valuation._TimeBasedValuation.length === 0) {
                                // Create empty time-based valuation if needed for proper structure
                                if (valuation.FixedAssetObjectActionCode !== "03") {
                                    valuation._TimeBasedValuation = [{
                                        CompanyCodeCurrency: this.config.defaultCurrency,
                                        FixedAssetObjectActionCode: valuation.FixedAssetObjectActionCode === "01" ? "01" : "02"
                                    }];
                                }
                            }
                        });
                    }
                });
            }

            // Validate payload against metadata structure
            this._validatePayloadStructure(payload);

            return payload;
        },
        _validatePayloadStructure: function (payload) {
            console.log("[DataTransformer] Validating payload structure against metadata");

            // Validate _AccountAssignment structure
            if (payload._AccountAssignment) {
                if (!Array.isArray(payload._AccountAssignment)) {
                    console.error("_AccountAssignment must be an array according to metadata");
                    throw new Error("Invalid _AccountAssignment structure - must be array");
                }

                payload._AccountAssignment.forEach((entry, index) => {
                    // Validate required FixedAssetObjectActionCode
                    if (!entry.FixedAssetObjectActionCode) {
                        console.warn(`_AccountAssignment[${index}] missing required FixedAssetObjectActionCode`);
                    }

                    // Validate WBSElementExternalID length (max 24 chars per metadata)
                    if (entry.WBSElementExternalID && entry.WBSElementExternalID.length > 24) {
                        console.error(`WBSElementExternalID exceeds maximum length of 24 characters: ${entry.WBSElementExternalID}`);
                        throw new Error(`WBSElementExternalID too long: ${entry.WBSElementExternalID.length} chars (max 24)`);
                    }

                    // Validate other fields based on metadata
                    this._validateAccountAssignmentFields(entry);
                });
            }

            // Validate _Ledger structure
            if (payload._Ledger) {
                if (!Array.isArray(payload._Ledger)) {
                    console.error("_Ledger must be an array according to metadata");
                    throw new Error("Invalid _Ledger structure - must be array");
                }
            }

            console.log("[DataTransformer] Payload structure validation completed");
        },

        // NEW: Validate AccountAssignment fields against metadata
        _validateAccountAssignmentFields: function (entry) {
            const fieldValidations = {
                CostCenter: { maxLength: 10 },
                WBSElementExternalID: { maxLength: 24 },
                ProfitCenter: { maxLength: 10 },
                Segment: { maxLength: 10 },
                Plant: { maxLength: 4 },
                AssetLocation: { maxLength: 10 },
                Room: { maxLength: 8 },
                VehicleLicensePlateNumber: { maxLength: 15 },
                TaxJurisdiction: { maxLength: 15 },
                BusinessPlace: { maxLength: 4 },
                Fund: { maxLength: 10 },
                GrantID: { maxLength: 20 },
                FunctionalArea: { maxLength: 16 }
            };

            Object.keys(fieldValidations).forEach(fieldName => {
                if (entry[fieldName]) {
                    const validation = fieldValidations[fieldName];
                    const value = String(entry[fieldName]);

                    if (validation.maxLength && value.length > validation.maxLength) {
                        console.error(`${fieldName} exceeds maximum length of ${validation.maxLength}: ${value} (${value.length} chars)`);
                        throw new Error(`${fieldName} too long: ${value.length} chars (max ${validation.maxLength})`);
                    }
                }
            });
        },
        _buildODataSectionWithLedgerArea: function (payload, section, fieldName, value, changeInfo, changeData) {
            // Use existing changeableFields configuration
            const fieldConfig = this.changeableFields[changeInfo.baseField || fieldName];

            console.log(`[DEBUG] Building OData section for field: ${fieldName}, section: ${section}, value: ${value}`);

            switch (section) {
                case '_General':
                    if (!payload._General) payload._General = {};
                    payload._General[fieldName] = value;
                    console.log(`[DEBUG] Added to _General: ${fieldName} = ${value}`);
                    break;

                case '_Root':
                    // Handle root-level parameters like YY1_WBSELEMENT1_FAA
                    payload[fieldName] = value;
                    console.log(`[DEBUG] Added as root parameter: ${fieldName} = ${value}`);
                    break;

                case '_AccountAssignment':
                    // FIXED: Create _AccountAssignment as ARRAY based on metadata
                    if (!payload._AccountAssignment) {
                        payload._AccountAssignment = [];
                    }

                    // Find existing entry or create new one
                    let accountAssignmentEntry = payload._AccountAssignment[0];
                    if (!accountAssignmentEntry) {
                        accountAssignmentEntry = {
                            // Set action code - use provided action code or default to "02" (Change)
                            FixedAssetObjectActionCode: changeData.actionCodes?.accountAssignment || this.config.defaultActionCode
                        };

                        // Extract ValidityStartDate from Excel if available
                        const validityStartDate = this._extractDateFromExcel(
                            changeData._originalExcelData,
                            'AccountAssignment_ValidityStartDate'
                        );
                        if (validityStartDate) {
                            accountAssignmentEntry.ValidityStartDate = validityStartDate;
                            console.log(`[DEBUG] Added ValidityStartDate: ${validityStartDate}`);
                        }

                        payload._AccountAssignment.push(accountAssignmentEntry);
                        console.log(`[DEBUG] Created _AccountAssignment array with action code: ${accountAssignmentEntry.FixedAssetObjectActionCode}`);
                    }

                    // Add the field to AccountAssignment entry
                    accountAssignmentEntry[fieldName] = value;
                    console.log(`[DEBUG] Added to _AccountAssignment[0]: ${fieldName} = ${value}`);
                    break;

                case '_GlobMasterData':
                    if (!payload._GlobMasterData) payload._GlobMasterData = {};
                    if (!payload._GlobMasterData._IN_AssetBlockData) {
                        payload._GlobMasterData._IN_AssetBlockData = {};
                    }
                    payload._GlobMasterData._IN_AssetBlockData[fieldName] = value;
                    break;

                case '_Inventory':
                    if (!payload._Inventory) payload._Inventory = {};
                    payload._Inventory[fieldName] = value;
                    break;

                case '_Ledger':
                    // Use enhanced Ledger_Area mapping logic
                    this._addLedgerFieldToPayloadWithLedgerArea(payload, fieldName, value, changeInfo, changeData);
                    break;
            }
        },

        /**
         * Add ledger field to payload with Ledger_Area format
         * CRITICAL FIX: Handle ValidityStartDate correctly for CHANGE vs CREATE actions
         * @private
         */
        _addLedgerFieldToPayloadWithLedgerArea: function (payload, fieldName, value, changeInfo, changeData) {
            const ledgerAreaMatch = fieldName.match(/^(.+?)_(\w+_\w+)$/);
            if (!ledgerAreaMatch) return;

            const [, baseField, ledgerAreaCode] = ledgerAreaMatch;

            const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerAreaCode];
            if (!ledgerAreaMapping) {
                console.warn(`[DataTransformer] Unknown Ledger_Area code: ${ledgerAreaCode}`);
                return;
            }

            const { ledger: ledgerCode, area: areaCode } = ledgerAreaMapping;

            const ledgerActionCode = changeData.actionCodes.ledgers[ledgerAreaCode] || this.config.defaultActionCode;
            const valuationActionCode = changeData.actionCodes.valuations[ledgerAreaCode] || this.config.defaultActionCode;

            // v3.0: use explicit TBV action code if provided, else derive from valuation action
            const timeBasedValuationActionCode =
                (changeData.actionCodes.timeBasedValuations && changeData.actionCodes.timeBasedValuations[ledgerAreaCode]) ||
                (valuationActionCode === "01" ? "01" : "02");

            console.log(`[DataTransformer] Processing field ${fieldName} for ${ledgerAreaCode}, actions: L=${ledgerActionCode}, V=${valuationActionCode}, TBV=${timeBasedValuationActionCode}`);

            if (!payload._Ledger) payload._Ledger = [];

            // Find or create ledger entry
            let ledgerEntry = payload._Ledger.find(l => l.Ledger === ledgerCode);
            if (!ledgerEntry) {
                ledgerEntry = {
                    Ledger: ledgerCode,
                    FixedAssetObjectActionCode: ledgerActionCode,
                    _Valuation: []
                };

                const assetCapitalizationDate =
                    this._extractDateFromExcel(changeData._originalExcelData, `AssetCapitalizationDate_${ledgerAreaCode}`) ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'AssetCapitalizationDate') ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'IN_AssetPutToUseDate');

                if (assetCapitalizationDate) {
                    ledgerEntry.AssetCapitalizationDate = assetCapitalizationDate;
                    console.log(`[DataTransformer] Set AssetCapitalizationDate: ${assetCapitalizationDate}`);
                }

                payload._Ledger.push(ledgerEntry);
            } else {
                if (ledgerActionCode && ledgerActionCode !== "02") {
                    ledgerEntry.FixedAssetObjectActionCode = ledgerActionCode;
                }
            }

            if (ledgerActionCode === "03") {
                console.log(`[DataTransformer] Ledger DELETE action - skipping valuation`);
                return;
            }

            // v3.0: AssetDeactivationDate is a ledger-level field
            if (baseField === 'AssetDeactivationDate') {
                ledgerEntry.AssetDeactivationDate = value;
                return;
            }

            // Ledger-level capitalization date override
            if (baseField === 'AssetCapitalizationDate') {
                ledgerEntry.AssetCapitalizationDate = value;
                return;
            }

            // Find or create valuation entry
            let valuationEntry = ledgerEntry._Valuation.find(v => v.AssetDepreciationArea === areaCode);
            if (!valuationEntry) {
                // v3.0: read NegativeAmountIsAllowed from Excel, fallback to true
                const rawNegativeAllowed =
                    changeData._originalExcelData &&
                    changeData._originalExcelData[`NegativeAmountIsAllowed_${ledgerAreaCode}`];
                const negativeAmountIsAllowed = (rawNegativeAllowed === null || rawNegativeAllowed === undefined || rawNegativeAllowed === "")
                    ? true
                    : this._parseBoolean(rawNegativeAllowed);

                valuationEntry = {
                    AssetDepreciationArea: areaCode,
                    FixedAssetObjectActionCode: valuationActionCode,
                    NegativeAmountIsAllowed: negativeAmountIsAllowed, // v3.0: from Excel, not hardcoded
                    _TimeBasedValuation: []
                };

                const depreciationStartDate =
                    this._extractDateFromExcel(changeData._originalExcelData, `DepreciationStartDate_${ledgerAreaCode}`) ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'DepreciationStartDate') ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'IN_AssetPutToUseDate') ||
                    this._extractDateFromExcel(changeData._originalExcelData, `AssetCapitalizationDate_${ledgerAreaCode}`) ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'AssetCapitalizationDate');

                if (depreciationStartDate) {
                    valuationEntry.DepreciationStartDate = depreciationStartDate;
                    console.log(`[DataTransformer] Set DepreciationStartDate: ${depreciationStartDate}`);
                }

                ledgerEntry._Valuation.push(valuationEntry);
            } else {
                valuationEntry.FixedAssetObjectActionCode = valuationActionCode;
            }

            if (valuationActionCode === "03") {
                console.log(`[DataTransformer] Valuation DELETE action - skipping time-based valuation`);
                return;
            }

            // v3.0: SpecialDeprStartDate and InterestCalcStartDate are valuation-level
            if (baseField === 'SpecialDeprStartDate') {
                valuationEntry.SpecialDeprStartDate = value;
                return;
            }
            if (baseField === 'InterestCalcStartDate') {
                valuationEntry.InterestCalcStartDate = value;
                return;
            }

            // Valuation-level depreciation start date override
            if (baseField === 'DepreciationStartDate') {
                valuationEntry.DepreciationStartDate = value;
                return;
            }

            // Find or create time-based valuation entry
            let timeBasedEntry = valuationEntry._TimeBasedValuation[0];
            if (!timeBasedEntry) {
                timeBasedEntry = {
                    CompanyCodeCurrency: this.config.defaultCurrency,
                    FixedAssetObjectActionCode: timeBasedValuationActionCode, // v3.0: uses TBV action code
                    ScrapAmountInCoCodeCrcy: 0
                };

                const explicitValidityDate = this._extractDateFromExcel(
                    changeData._originalExcelData,
                    `ValidityStartDate_${ledgerAreaCode}`
                );

                if (timeBasedValuationActionCode === "01") {
                    // CREATE: must have ValidityStartDate
                    const validityStartDate =
                        explicitValidityDate ||
                        this._extractDateFromExcel(changeData._originalExcelData, 'ValidityStartDate') ||
                        this._extractDateFromExcel(changeData._originalExcelData, 'IN_AssetPutToUseDate') ||
                        this._extractDateFromExcel(changeData._originalExcelData, `DepreciationStartDate_${ledgerAreaCode}`) ||
                        this._extractDateFromExcel(changeData._originalExcelData, 'DepreciationStartDate');

                    if (validityStartDate) {
                        timeBasedEntry.ValidityStartDate = validityStartDate;
                        console.log(`[DataTransformer] CREATE action - Set ValidityStartDate: ${validityStartDate}`);
                    } else {
                        console.warn(`[DataTransformer] CREATE action but no ValidityStartDate found for ${ledgerAreaCode}`);
                    }
                } else if (timeBasedValuationActionCode === "02" && explicitValidityDate) {
                    // CHANGE: only add if user explicitly provided it
                    timeBasedEntry.ValidityStartDate = explicitValidityDate;
                    console.log(`[DataTransformer] CHANGE action - Using explicit ValidityStartDate: ${explicitValidityDate}`);
                } else {
                    console.log(`[DataTransformer] CHANGE action - No ValidityStartDate specified, SAP will use existing date`);
                }

                valuationEntry._TimeBasedValuation.push(timeBasedEntry);
            }

            // All remaining fields go into _TimeBasedValuation
            this._setTimeBasedValuationField(timeBasedEntry, baseField, value);
        },
        /**
         * Validate ValidityStartDate requirements for time-based valuation
         * @private
         */
        _validateValidityStartDateRequirements: function (changeData, errors, addUniqueError) {
            const sequenceId = changeData.SequenceNumber || 'N/A';
            const asset = changeData.FixedAsset || 'Unknown';

            Object.keys(this.config.ledgerAreaToLedgerMap).forEach(ledgerArea => {
                const valuationAction = changeData.actionCodes.valuations?.[ledgerArea];
                const tbvAction = changeData.actionCodes.timeBasedValuations?.[ledgerArea]; // v3.0 NEW

                // Only validate if there is a valuation or TBV action
                if (!valuationAction && !tbvAction) return;

                const explicitValidityDate = this._extractDateFromExcel(
                    changeData._originalExcelData,
                    `ValidityStartDate_${ledgerArea}`
                );

                const hasFallbackDate =
                    explicitValidityDate ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'ValidityStartDate') ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'IN_AssetPutToUseDate') ||
                    this._extractDateFromExcel(changeData._originalExcelData, `DepreciationStartDate_${ledgerArea}`) ||
                    this._extractDateFromExcel(changeData._originalExcelData, 'DepreciationStartDate');

                // ── Valuation-level rules ──────────────────────────────────────────

                if (valuationAction === "01") {
                    if (!hasFallbackDate) {
                        addUniqueError({
                            field: `ValidityStartDate_${ledgerArea}`,
                            message: `ValidityStartDate is REQUIRED when creating new depreciation for ${ledgerArea}. Provide ValidityStartDate_${ledgerArea}_Value or a general ValidityStartDate_Value.`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'MANDATORY_FIELD'
                        });
                    }
                }

                if (valuationAction === "02" && explicitValidityDate) {
                    addUniqueError({
                        field: `ValidityStartDate_${ledgerArea}`,
                        message: `WARNING: When using CHANGE action (02), ValidityStartDate must match an EXISTING date in SAP. Date provided: ${explicitValidityDate}. If this date doesn't exist in SAP, you'll get error: "Dep. area ${ledgerArea.split('_')[1]} of value base time interval starting with ${explicitValidityDate} doesn't exist". To avoid this error, leave ValidityStartDate blank and SAP will use the existing date.`,
                        severity: 'Warning',
                        sequenceId: sequenceId,
                        asset: asset,
                        validationType: 'VALIDITY_DATE_WARNING'
                    });
                }

                // ── v3.0 NEW: TimeBasedValuation-level rules (independent) ────────
                // Only validate when TBV action differs from Valuation action —
                // same action is already covered by the Valuation-level rules above.
                if (tbvAction && tbvAction !== valuationAction) {
                    if (tbvAction === "01") {
                        if (!hasFallbackDate) {
                            addUniqueError({
                                field: `ValidityStartDate_${ledgerArea}`,
                                message: `ValidityStartDate is REQUIRED when TimeBasedValuation_ActionCode is CREATE (01) for ${ledgerArea}, even if Valuation action is CHANGE. Provide ValidityStartDate_${ledgerArea}_Value.`,
                                severity: 'Error',
                                sequenceId: sequenceId,
                                asset: asset,
                                validationType: 'MANDATORY_FIELD'
                            });
                        }
                    } else if (tbvAction === "02" && explicitValidityDate) {
                        addUniqueError({
                            field: `ValidityStartDate_${ledgerArea}`,
                            message: `WARNING: TimeBasedValuation_ActionCode CHANGE (02) for ${ledgerArea} — ValidityStartDate (${explicitValidityDate}) must EXACTLY match an existing interval in SAP, or leave blank to use the existing date.`,
                            severity: 'Warning',
                            sequenceId: sequenceId,
                            asset: asset,
                            validationType: 'VALIDITY_DATE_WARNING'
                        });
                    }
                }
            });
        },

        /**
         * Set field value in time-based valuation with correct data types
         * @private
         */
        _setTimeBasedValuationField: function (timeBasedEntry, baseField, value) {
            switch (baseField) {
                case 'DepreciationKey':
                    // String field
                    timeBasedEntry.DepreciationKey = value;
                    break;

                case 'PlannedUsefulLifeInYears':
                    // String field (SAP metadata shows "string")
                    timeBasedEntry.PlannedUsefulLifeInYears = String(value);
                    break;

                case 'PlannedUsefulLifeInPeriods':
                    // String field (SAP metadata shows "string")
                    timeBasedEntry.PlannedUsefulLifeInPeriods = String(value);
                    break;

                case 'ScrapAmountInCoCodeCrcy':
                    // NUMERIC field (SAP metadata shows 0, not "0")
                    timeBasedEntry.ScrapAmountInCoCodeCrcy = parseFloat(value) || 0;
                    break;

                case 'AcqnProdnCostScrapPercent':
                    // NUMERIC field (SAP metadata shows 0, not "0")
                    timeBasedEntry.AcqnProdnCostScrapPercent = parseFloat(value) || 0;
                    break;

                case 'VariableDeprPercent':
                    // NUMERIC field
                    timeBasedEntry.VariableDeprPercent = parseFloat(value) || 0;
                    break;

                case 'DeprCalcBaseValuePercent':
                    // NUMERIC field
                    timeBasedEntry.DeprCalcBaseValuePercent = parseFloat(value) || 0;
                    break;

                case 'ShiftOperationFactor':
                    // NUMERIC field
                    timeBasedEntry.ShiftOperationFactor = parseFloat(value) || 0;
                    break;

                case 'IsShutDown':
                    // Boolean field
                    timeBasedEntry.IsShutDown = this._parseBoolean(value);
                    break;

                // Scrap amounts for different currencies (all numeric)
                case 'ScrapAmountInGlobCrcy':
                    timeBasedEntry.ScrapAmountInGlobCrcy = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy1':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy1 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy2':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy2 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy3':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy3 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy4':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy4 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy5':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy5 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy6':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy6 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy7':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy7 = parseFloat(value) || 0;
                    break;
                case 'ScrapAmountInFreeDefinedCrcy8':
                    timeBasedEntry.ScrapAmountInFreeDefinedCrcy8 = parseFloat(value) || 0;
                    break;

                // Asset usage object for unit-of-production depreciation
                case 'FixedAssetUsageObject':
                    timeBasedEntry.FixedAssetUsageObject = value;
                    break;

                // Asset revaluation index
                case 'AssetRevaluationIndex':
                    timeBasedEntry.AssetRevaluationIndex = value;
                    break;
            }
        },

        // =================
        // DISPLAY PROCESSING WITH LEDGER_AREA FORMAT
        // =================

        processChangeEntryDataForDisplay: function (entry) {
            const processed = JSON.parse(JSON.stringify(entry));

            processed.ProcessedValidationErrors = this._processValidationErrors(processed.ValidationErrors || []);
            processed.ChangesSummary = this._processChangesForDisplayWithLedgerArea(processed.changes || {});
            processed.Status = processed.Status || (processed.ProcessedValidationErrors.length > 0 ? 'Invalid' : 'Valid');
            processed.HasBlankAssignments = this._hasBlankAssignments(processed.changes || {});
            processed.AllFields = this._createAllFieldsArrayWithLedgerArea(processed);
            processed.ActionCodes = entry.actionCodes || {}; // Include action codes

            // Ensure validation errors have asset information
            if (processed.ProcessedValidationErrors && processed.ProcessedValidationErrors.length > 0) {
                processed.ProcessedValidationErrors.forEach(error => {
                    if (!error.asset) {
                        error.asset = processed.FixedAsset || 'Unknown';
                    }
                });
            }

            return processed;
        },

        _processChangesForDisplayWithLedgerArea: function (changes) {
            const summary = {
                totalChanges: 0,
                blankAssignments: 0,
                fieldChanges: [],
                sectionChanges: {},
                actionSummary: {},
                ledgerActions: {},
                valuationActions: {},
                timeBasedValuationActions: {} // v3.0 NEW
            };

            Object.entries(changes).forEach(([fieldName, changeInfo]) => {
                if (changeInfo.shouldChange) {
                    summary.totalChanges++;

                    if (changeInfo.isBlankValue) {
                        summary.blankAssignments++;
                    }

                    let actionCode = this.config.defaultActionCode;
                    let actionName = "Change";

                    if (changeInfo.ledgerActionCode) {
                        actionCode = changeInfo.ledgerActionCode;
                        actionName = this._getActionName(actionCode);
                        if (changeInfo.ledgerAreaCode) {
                            summary.ledgerActions[changeInfo.ledgerAreaCode] = actionCode;
                        }
                    }

                    if (changeInfo.valuationActionCode) {
                        actionCode = changeInfo.valuationActionCode;
                        actionName = this._getActionName(actionCode);
                        if (changeInfo.ledgerAreaCode) {
                            summary.valuationActions[changeInfo.ledgerAreaCode] = actionCode;
                        }
                    }

                    // v3.0 NEW: track TBV action independently
                    if (changeInfo.timeBasedValuationActionCode && changeInfo.ledgerAreaCode) {
                        summary.timeBasedValuationActions[changeInfo.ledgerAreaCode] =
                            changeInfo.timeBasedValuationActionCode;
                    }

                    const displayInfo = {
                        fieldName: fieldName,
                        displayName: this._getFieldDisplayNameWithLedgerArea(fieldName),
                        newValue: changeInfo.isBlankValue ? "<Blank>" : changeInfo.newValue,
                        isBlankValue: changeInfo.isBlankValue,
                        section: this._getFieldSection(fieldName),
                        actionCode: actionCode,
                        actionName: actionName,
                        ledgerAreaCode: changeInfo.ledgerAreaCode,
                        areaCode: changeInfo.areaCode,
                        timeBasedValuationActionCode: changeInfo.timeBasedValuationActionCode || null // v3.0 NEW
                    };

                    summary.fieldChanges.push(displayInfo);

                    if (!summary.sectionChanges[displayInfo.section]) {
                        summary.sectionChanges[displayInfo.section] = [];
                    }
                    summary.sectionChanges[displayInfo.section].push(displayInfo);

                    if (!summary.actionSummary[actionName]) {
                        summary.actionSummary[actionName] = 0;
                    }
                    summary.actionSummary[actionName]++;
                }
            });

            summary.createCount = summary.actionSummary['Create'] || 0;
            summary.changeCount = summary.actionSummary['Change'] || 0;
            summary.deleteCount = summary.actionSummary['Delete'] || 0;

            return summary;
        },

        _getActionName: function (actionCode) {
            switch (actionCode) {
                case "01": return "Create";
                case "02": return "Change";
                case "03": return "Delete";
                default: return "Change";
            }
        },

        _calculateChangesSummary: function (changes) {
            const summary = {
                totalChanges: 0,
                blankAssignments: 0,
                fieldChanges: [],
                createCount: 0,
                changeCount: 0,
                deleteCount: 0,
                ledgerActions: {},
                valuationActions: {},
                timeBasedValuationActions: {} // v3.0 NEW
            };

            Object.entries(changes).forEach(([fieldName, changeInfo]) => {
                if (changeInfo && changeInfo.shouldChange) {
                    summary.totalChanges++;

                    if (changeInfo.isBlankValue) {
                        summary.blankAssignments++;
                    }

                    // Determine the effective action code
                    let actionCode = this.config.defaultActionCode;

                    if (changeInfo.ledgerActionCode) {
                        actionCode = changeInfo.ledgerActionCode;
                        if (changeInfo.ledgerAreaCode) {
                            summary.ledgerActions[changeInfo.ledgerAreaCode] = actionCode;
                        }
                    }

                    if (changeInfo.valuationActionCode) {
                        actionCode = changeInfo.valuationActionCode;
                        if (changeInfo.ledgerAreaCode) {
                            summary.valuationActions[changeInfo.ledgerAreaCode] = actionCode;
                        }
                    }

                    // v3.0 NEW: track TBV action independently
                    if (changeInfo.timeBasedValuationActionCode && changeInfo.ledgerAreaCode) {
                        summary.timeBasedValuationActions[changeInfo.ledgerAreaCode] =
                            changeInfo.timeBasedValuationActionCode;
                    }

                    if (actionCode === "01") summary.createCount++;
                    else if (actionCode === "02") summary.changeCount++;
                    else if (actionCode === "03") summary.deleteCount++;

                    summary.fieldChanges.push({
                        field: fieldName,
                        newValue: changeInfo.newValue,
                        isBlankValue: changeInfo.isBlankValue,
                        actionCode: actionCode,
                        actionName: this._getActionName(actionCode),
                        ledgerAreaCode: changeInfo.ledgerAreaCode,
                        areaCode: changeInfo.areaCode,
                        timeBasedValuationActionCode: changeInfo.timeBasedValuationActionCode || null // v3.0 NEW
                    });
                }
            });

            return summary;
        },

        _createAllFieldsArrayWithLedgerArea: function (entry) {
            const allFields = [];
            const changes = entry.changes || {};

            this.FIELD_DEFINITIONS.forEach(fieldDef => {
                const changeInfo = changes[fieldDef.key];
                const isChanged = changeInfo && changeInfo.shouldChange;
                const currentValue = this._getCurrentFieldValue(entry, fieldDef.key);

                let actionCode = null;
                let actionName = null;

                if (isChanged) {
                    if (fieldDef.key.startsWith('Ledger_ActionCode_') || fieldDef.key.startsWith('Valuation_ActionCode_')) {
                        actionCode = currentValue;
                        actionName = this._getActionName(actionCode);
                    } else if (changeInfo.ledgerActionCode || changeInfo.valuationActionCode) {
                        actionCode = changeInfo.valuationActionCode || changeInfo.ledgerActionCode || this.config.defaultActionCode;
                        actionName = this._getActionName(actionCode);
                    } else {
                        actionCode = this.config.defaultActionCode;
                        actionName = this._getActionName(actionCode);
                    }
                }

                allFields.push({
                    fieldName: fieldDef.key,
                    displayName: fieldDef.displayName,
                    section: fieldDef.section,
                    isChanged: isChanged,
                    currentValue: currentValue || '',
                    newValue: isChanged ? changeInfo.newValue : currentValue,
                    isBlank: isChanged ? changeInfo.isBlankValue : false,
                    actionCode: actionCode,
                    actionName: actionName,
                    ledgerAreaCode: isChanged ? changeInfo.ledgerAreaCode : null
                });
            });

            return allFields;
        },

        _getCurrentFieldValue: function (entry, fieldKey) {
            // Check for _Value suffix first
            const valueKey = fieldKey + '_Value';
            if (entry.hasOwnProperty(valueKey)) {
                return entry[valueKey];
            }

            // Check direct field name
            if (entry.hasOwnProperty(fieldKey)) {
                return entry[fieldKey];
            }

            // Check in processed sections
            const sections = ['GeneralDetails', 'AccountAssignment', 'GlobalMasterData', 'InventoryData'];
            for (const section of sections) {
                if (entry[section] && entry[section][fieldKey] && entry[section][fieldKey].Value !== undefined) {
                    return entry[section][fieldKey].Value;
                }
            }

            // Check depreciation areas with Ledger_Area format
            if (entry.DepreciationAreas && Array.isArray(entry.DepreciationAreas)) {
                const ledgerAreaMatch = fieldKey.match(/^(.+?)_(\w+_\w+)$/);
                if (ledgerAreaMatch) {
                    const [, baseField, ledgerAreaCode] = ledgerAreaMatch;
                    const ledgerAreaMapping = this.config.ledgerAreaToLedgerMap[ledgerAreaCode];
                    if (ledgerAreaMapping) {
                        const area = entry.DepreciationAreas.find(a => a.DepreciationArea === ledgerAreaMapping.area);
                        if (area && area[baseField] && area[baseField].Value !== undefined) {
                            return area[baseField].Value;
                        }
                    }
                }
            }

            return '';
        },

        processValidationErrorsForDisplay: function (errors) {
            if (!errors || !Array.isArray(errors) || errors.length === 0) {
                return {
                    flatErrors: [],
                    totalErrors: 0,
                    affectedAssets: 0
                };
            }

            const processedErrors = errors
                .filter(error => {
                    if (typeof error === 'string') return true;
                    const severity = error.severity || error.Severity || 'Error';
                    return severity === 'Error';
                })
                .map(error => {
                    if (typeof error === 'string') {
                        return {
                            message: error,
                            field: 'Unknown',
                            severity: 'Error',
                            sequenceId: 'N/A',
                            asset: 'Unknown'
                        };
                    }

                    return {
                        message: error.message || error.Message || JSON.stringify(error),
                        field: error.field || error.Field || 'Unknown',
                        severity: 'Error',
                        sequenceId: error.sequenceId || error.SequenceId || 'N/A',
                        asset: error.asset || error.Asset || 'Unknown'
                    };
                });

            const uniqueAssets = new Set();
            processedErrors.forEach(error => {
                if (error.asset && error.asset !== 'Unknown') {
                    uniqueAssets.add(error.asset);
                }
            });

            return {
                flatErrors: processedErrors,
                totalErrors: processedErrors.length,
                affectedAssets: uniqueAssets.size
            };
        },

        calculateValidationStats: function (total, successful, failed) {
            const successRate = total > 0 ? Math.round((successful / total) * 100) : 0;
            const failureRate = total > 0 ? Math.round((failed / total) * 100) : 0;

            return {
                TotalAssets: total,
                ValidAssets: successful,
                InvalidAssets: failed,
                SuccessRate: successRate,
                FailureRate: failureRate,
                ProcessingComplete: true,
                IsValid: failed === 0 && total > 0,
                Status: failed === 0 ? 'Valid' : 'Invalid',
                Stats: [
                    { name: "Valid", value: successful, percent: successRate },
                    { name: "Invalid", value: failed, percent: failureRate }
                ]
            };
        },

        _hasBlankAssignments: function (changes) {
            return Object.values(changes).some(change =>
                change.shouldChange && change.isBlankValue
            );
        },

        _getFieldDisplayNameWithLedgerArea: function (fieldName) {
            // Check for Ledger_Area specific fields
            const ledgerAreaMatch = fieldName.match(/^(.+?)_(\w+_\w+)$/);
            if (ledgerAreaMatch) {
                const [, baseField, ledgerAreaCode] = ledgerAreaMatch;
                const baseDisplayName = this.DISPLAY_NAMES[baseField] || baseField;
                return `${baseDisplayName} (${ledgerAreaCode})`;
            }

            return this.DISPLAY_NAMES[fieldName] || fieldName;
        },

        _getFieldSection: function (fieldName) {
            // Special handling for YY1_WBSELEMENT1_FAA - it's a root parameter
            if (fieldName === 'YY1_WBSELEMENT1_FAA') {
                return '_Root';
            }

            // Check for Ledger_Area specific fields
            const ledgerAreaMatch = fieldName.match(/^(.+?)_(\w+_\w+)$/);
            if (ledgerAreaMatch) {
                const [, baseField, ledgerAreaCode] = ledgerAreaMatch;

                // Only treat as ledger field if it's actually a known ledger area
                if (this.config.ledgerAreaToLedgerMap[ledgerAreaCode]) {
                    const baseFieldConfig = this.changeableFields[baseField];
                    if (baseFieldConfig) {
                        return baseFieldConfig.section;
                    }
                    return '_Ledger';
                }
                // If not a known ledger area, fall through to normal processing
            }

            const fieldConfig = this.changeableFields[fieldName];
            if (fieldConfig) {
                return fieldConfig.section;
            }

            if (fieldName.includes('_')) {
                return '_Ledger';
            }

            return '_General';
        },

        processRowsForExport: function (data, reportType = "all") {
            if (!data || !Array.isArray(data)) {
                console.warn("DataTransformer: No data provided for export processing");
                return [];
            }

            let filteredData = data;

            // Filter data based on report type
            if (reportType === "success") {
                filteredData = data.filter(item =>
                    item.Status === "Success" ||
                    item.Status === "Valid" ||
                    item.success === true
                );
            } else if (reportType === "error" || reportType === "errors") {
                filteredData = data.filter(item =>
                    item.Status === "Error" ||
                    item.Status === "Failed" ||
                    item.Status === "Invalid" ||
                    item.success === false
                );
            }

            // Process each row for export with Ledger_Area format
            return filteredData.map((item, index) => {
                const processedItem = { ...item };

                // Ensure basic fields
                processedItem.SequenceNumber = processedItem.SequenceNumber ||
                    processedItem.Sequence ||
                    (index + 1);

                // Process change summary if available
                if (item.ChangesSummary) {
                    processedItem.ChangesCount = item.ChangesSummary.totalChanges || 0;
                    processedItem.BlankAssignments = item.ChangesSummary.blankAssignments || 0;
                    processedItem.CreateCount = item.ChangesSummary.createCount || 0;
                    processedItem.ChangeCount = item.ChangesSummary.changeCount || 0;
                    processedItem.DeleteCount = item.ChangesSummary.deleteCount || 0;

                    // Process ledger actions with Ledger_Area format
                    if (item.ChangesSummary.ledgerActions) {
                        Object.entries(item.ChangesSummary.ledgerActions).forEach(([ledgerArea, action]) => {
                            processedItem[`LedgerAction_${ledgerArea}`] = this._getActionName(action);
                        });
                    }

                    // Process valuation actions with Ledger_Area format
                    if (item.ChangesSummary.valuationActions) {
                        Object.entries(item.ChangesSummary.valuationActions).forEach(([ledgerArea, action]) => {
                            processedItem[`ValuationAction_${ledgerArea}`] = this._getActionName(action);
                        });
                    }

                    if (item.ChangesSummary.fieldChanges) {
                        processedItem.ChangedFields = item.ChangesSummary.fieldChanges
                            .map(fc => `${fc.displayName || fc.fieldName} (${fc.actionName})`)
                            .join(", ");
                    }
                }

                // Process validation errors
                if (item.ValidationErrors && Array.isArray(item.ValidationErrors)) {
                    processedItem.ValidationErrorCount = item.ValidationErrors.length;
                    processedItem.ValidationErrorSummary = item.ValidationErrors
                        .map(err => err.message || err.Message)
                        .join("; ");
                }

                // Clean up internal fields
                const fieldsToRemove = [
                    '_sheetName', '_originalData', '_originalIndex',
                    'ValidationErrors', 'ProcessedValidationErrors',
                    'ChangesSummary', '@odata.context', '@odata.etag',
                    '_originalExcelData' // Remove the Excel data reference
                ];

                fieldsToRemove.forEach(field => delete processedItem[field]);

                return processedItem;
            });
        },

        /**
         * Process batch results for export
         * @public
         */
        processBatchResults: function (batchResultsData, reportType = "all") {
            console.log("[DataTransformer] Processing batch results for export");
            console.log("[DataTransformer] Input data:", batchResultsData);
            console.log("[DataTransformer] Report type:", reportType);

            if (!batchResultsData) {
                console.warn("[DataTransformer] No batch results data provided");
                return {
                    processedRecords: [],
                    summary: { totalRecords: 0, successCount: 0, errorCount: 0, reportType }
                };
            }

            // Handle different input formats
            let dataArray = [];
            if (Array.isArray(batchResultsData)) {
                dataArray = batchResultsData;
            } else if (batchResultsData.successRecords || batchResultsData.errorRecords || batchResultsData.results) {
                // Handle wrapped data formats
                dataArray = [
                    ...(batchResultsData.successRecords || []),
                    ...(batchResultsData.errorRecords || []),
                    ...(batchResultsData.results || [])
                ];
            } else {
                console.warn("[DataTransformer] Unrecognized data format");
                return {
                    processedRecords: [],
                    summary: { totalRecords: 0, successCount: 0, errorCount: 0, reportType }
                };
            }

            // Separate success and error records
            const successRecords = dataArray.filter(record => record["@odata.context"]);
            const errorRecords = dataArray.filter(record => record.error);

            console.log("[DataTransformer] Success records found:", successRecords.length);
            console.log("[DataTransformer] Error records found:", errorRecords.length);

            // Select records based on report type
            let recordsToProcess = [];
            switch (reportType) {
                case "success":
                    recordsToProcess = successRecords;
                    break;
                case "error":
                case "errors":
                    recordsToProcess = errorRecords;
                    break;
                default:
                    recordsToProcess = [...successRecords, ...errorRecords];
                    break;
            }

            console.log("[DataTransformer] Records to process:", recordsToProcess.length);

            // Process each record
            const processedRecords = recordsToProcess.map((record, index) => {
                const isSuccess = !!record["@odata.context"];

                if (isSuccess) {
                    // Handle success records
                    const originalRequest = record.OriginalRequest || {};
                    return {
                        SequenceNumber: originalRequest.SequenceNumber || (index + 1),
                        CompanyCode: record.CompanyCode || "",
                        FixedAsset: record.MasterFixedAsset || record.FixedAsset || "",
                        Status: "Success",
                        Message: record.Message || "Asset processed successfully",

                        // Asset details
                        AssetDescription: record.FixedAssetDescription || "",
                        SerialNumber: record.AssetSerialNumber || "",
                        InventoryNote: record.InventoryNote || "",

                        // Changes summary
                        ChangesSummary: originalRequest.ChangesSummary || {},
                        ChangedFieldsSummary: originalRequest.ChangesSummary?.fieldChanges?.map(fc =>
                            `${fc.field} (${fc.actionName || "Change"})`).join(", ") || "",

                        // For PDF formatter compatibility
                        ProcessedValidationErrors: [{
                            message: record.Message || "Asset processed successfully",
                            field: "Status",
                            severity: "Success",
                            sequenceId: originalRequest.SequenceNumber || (index + 1),
                            asset: record.MasterFixedAsset || record.FixedAsset || ""
                        }],

                        HasBlankAssignments: originalRequest.ChangesSummary?.blankAssignments > 0 || false,
                        RecordType: reportType
                    };
                } else {
                    // Handle error records
                    const entry = record.entry || {};
                    const errorMessage = record.details?.lastError || record.error || "Processing failed";
                    const errorDetails = record.details ? JSON.stringify(record.details) : "";

                    return {
                        SequenceNumber: entry.SequenceNumber || (index + 1),
                        CompanyCode: entry.CompanyCode || "",
                        FixedAsset: entry.FixedAsset || "",
                        Status: "Error",
                        ErrorMessage: errorMessage,
                        ErrorDetails: errorDetails,
                        ErrorCode: record.errorCode || "PROCESSING_ERROR",

                        // Asset details (empty for errors)
                        AssetDescription: "",
                        SerialNumber: "",
                        InventoryNote: "",

                        // Changes summary from entry
                        ChangesSummary: entry.ChangesSummary || {},
                        ChangedFieldsSummary: entry.ChangesSummary?.fieldChanges?.map(fc =>
                            `${fc.field} (${fc.actionName || "Change"})`).join(", ") || "",

                        // For PDF formatter compatibility
                        ProcessedValidationErrors: [{
                            message: errorMessage,
                            field: "Error",
                            severity: "Error",
                            sequenceId: entry.SequenceNumber || (index + 1),
                            asset: entry.FixedAsset || ""
                        }],

                        ValidationErrorsSummary: errorMessage,
                        HasBlankAssignments: entry.ChangesSummary?.blankAssignments > 0 || false,

                        ExportTimestamp: new Date().toISOString(),
                        RecordType: reportType
                    };
                }
            });

            const result = {
                processedRecords,
                summary: {
                    totalRecords: recordsToProcess.length,
                    successCount: successRecords.length,
                    errorCount: errorRecords.length,
                    reportType
                }
            };

            console.log("[DataTransformer] Final result:", result);
            return result;
        },

        // =================
        // UTILITY METHODS
        // =================

        parseDate: function (dateValue) {
            if (!dateValue) return null;

            try {
                const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
                return !isNaN(date.getTime()) ? date : null;
            } catch (error) {
                console.warn("Date parsing failed:", dateValue, error);
                return null;
            }
        },

        _getValue: function (obj, key, defaultValue = '') {
            if (!obj || typeof obj !== 'object') return defaultValue;
            const value = obj[key];
            return (value !== null && value !== undefined) ? value : defaultValue;
        },

        _parseBoolean: function (value) {
            if (typeof value === 'boolean') return value;
            if (typeof value === 'string') {
                const normalized = value.toLowerCase().trim();
                return this.config.changeFlagValues.positive.some(flag =>
                    String(flag).toLowerCase() === normalized
                );
            }
            return Boolean(value);
        },

        _processValidationErrors: function (errors) {
            return errors.map(error => {
                if (typeof error === 'string') {
                    return { message: error, field: 'Unknown', severity: 'Error' };
                }
                return {
                    message: error.message || error.Message || JSON.stringify(error),
                    field: error.field || error.Field || 'Unknown',
                    severity: error.severity || error.Severity || 'Error',
                    sequenceId: error.sequenceId || error.SequenceId || 'N/A'
                };
            }).filter(error => error.message);
        },

        _processChangeValue: function (value) {
            if (value === null || value === undefined) {
                return { value: "", isBlank: true };
            }

            const stringValue = String(value).trim();
            const isBlankMarker = this.config.blankMarkers.some(marker =>
                stringValue.toLowerCase() === marker.toLowerCase()
            );

            if (isBlankMarker || stringValue === "") {
                return { value: "", isBlank: true };
            }

            return { value: stringValue, isBlank: false };
        },

        _parseChangeFlag: function (value) {
            if (value === null || value === undefined) return false;

            const stringValue = String(value).trim().toUpperCase();
            return this.config.changeFlagValues.positive.some(flag =>
                String(flag).toUpperCase() === stringValue
            );
        },

        _validateFieldValue: function (fieldName, value, errors, sequenceId, asset) {
            if (!value || value === "") return;

            // Validation rules by field type
            const validationRules = {
                description: { maxLength: 50, fields: ['FixedAssetDescription', 'AssetAdditionalDescription'] },
                serialNumber: { maxLength: 18, fields: ['AssetSerialNumber'] },
                years: { min: 0, max: 999, fields: ['PlannedUsefulLifeInYears'] },
                periods: { min: 0, max: 9999, fields: ['PlannedUsefulLifeInPeriods'] },
                amount: { min: 0, fields: ['ScrapAmountInCoCodeCrcy'] },
                percentage: { min: 0, max: 100, fields: ['AcqnProdnCostScrapPercent'] },
                depreciationKey: { exactLength: 4, fields: ['DepreciationKey'] },
                companyCode: { exactLength: 4, fields: ['CompanyCode'] },
                date: { fields: ['IN_AssetPutToUseDate', 'AccountAssignment_ValidityStartDate', 'AssetCapitalizationDate', 'DepreciationStartDate', 'ValidityStartDate'] },
                wbs: { maxLength: 24, fields: ['WBSElementExternalID', 'YY1_WBSELEMENT1_FAA'] },
                room: { maxLength: 8, fields: ['Room'] },
                costCenter: { maxLength: 10, fields: ['CostCenter'] },
                profitCenter: { maxLength: 10, fields: ['ProfitCenter'] },
                plant: { maxLength: 4, fields: ['Plant'] },
                inventoryNote: { maxLength: 50, fields: ['InventoryNote'] }
            };

            // Find matching rule
            const rule = Object.values(validationRules).find(rule =>
                rule.fields.some(field => fieldName === field || fieldName.startsWith(field + '_'))
            );

            if (!rule) return;

            // Apply validation based on rule type
            if (rule.maxLength && value.length > rule.maxLength) {
                errors.push({
                    field: fieldName,
                    message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} cannot exceed ${rule.maxLength} characters (current: ${value.length})`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset
                });
            }

            if (rule.exactLength && value.length !== rule.exactLength) {
                errors.push({
                    field: fieldName,
                    message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be exactly ${rule.exactLength} characters`,
                    severity: 'Error',
                    sequenceId: sequenceId,
                    asset: asset
                });
            }

            if (rule.min !== undefined || rule.max !== undefined) {
                const numValue = parseFloat(value);
                if (isNaN(numValue)) {
                    errors.push({
                        field: fieldName,
                        message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be a valid number`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                } else {
                    if (rule.min !== undefined && numValue < rule.min) {
                        errors.push({
                            field: fieldName,
                            message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be at least ${rule.min}`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset
                        });
                    }
                    if (rule.max !== undefined && numValue > rule.max) {
                        errors.push({
                            field: fieldName,
                            message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be at most ${rule.max}`,
                            severity: 'Error',
                            sequenceId: sequenceId,
                            asset: asset
                        });
                    }
                }
            }

            // Date validation (only apply if this is the 'date' rule)
            if (rule === validationRules.date) {
                const date = this.parseDate(value);
                if (!date) {
                    errors.push({
                        field: fieldName,
                        message: `${this._getFieldDisplayNameWithLedgerArea(fieldName)} must be a valid date format (YYYY-MM-DD)`,
                        severity: 'Error',
                        sequenceId: sequenceId,
                        asset: asset
                    });
                }
            }
        },

        _mapBasicFields: function (row) {
            const mapped = {};

            Object.entries(row).forEach(([key, value]) => {
                const normalizedKey = key.toLowerCase().replace(/\s+/g, '');
                const mappedField = this.fieldMappings[normalizedKey];

                if (mappedField) {
                    mapped[mappedField] = value;
                } else {
                    mapped[key] = value;
                }
            });

            return mapped;
        }
    });
});