// Modified ExcelProcessor.js with Ledger_Area format support and ValidityStartDate validation
// v3.0: Dynamic LEDGER_AREAS + LEDGER_ACTION_CODE_LEDGERS, TimeBasedValuation_ActionCode support,
//       NegativeAmountIsAllowed / IsShutDown boolean field validation, dynamic ledgerAreaSummary

sap.ui.define([
  "sap/ui/base/Object",
  "sap/ui/model/json/JSONModel",
  "sap/m/MessageBox",
  "sap/m/MessageToast"
], function (
  BaseObject,
  JSONModel,
  MessageBox,
  MessageToast
) {
  "use strict";

  // ─── v3.0 Module-level constants ────────────────────────────────────────────

  /**
   * Default ledger areas — matches template generator default.
   * Override via options.ledgerAreas in the constructor.
   */
  const DEFAULT_LEDGER_AREAS = ["0L_01", "0L_15", "2L_32", "3L_34"];

  /**
   * Only these ledger codes carry Ledger_ActionCode at ledger level.
   * 2L and 3L have no FixedAssetObjectActionCode at ledger level per EDMX.
   * Override via options.ledgerActionCodeLedgers.
   */
  const DEFAULT_LEDGER_ACTION_CODE_LEDGERS = ["0L"];

  /**
   * Boolean-only fields — value must be true/false; no _Change pair.
   * Validated separately from standard Value/Change pairs.
   */
  const BOOLEAN_ONLY_FIELDS = ["NegativeAmountIsAllowed", "IsShutDown"];

  // ────────────────────────────────────────────────────────────────────────────

  return BaseObject.extend("assetmasterchange.service.ExcelProcessor", {

    constructor: function (options = {}) {
      if (!options.controller) {
        throw new Error("ExcelProcessor requires a controller instance");
      }
      if (!options.dataTransformer) {
        throw new Error("ExcelProcessor requires a shared DataTransformer instance");
      }

      this.oController = options.controller;
      this._dataTransformer = options.dataTransformer;
      this._uiManager = options.uiManager;

      // v3.0: configurable ledger area lists (backward-compat defaults)
      this.LEDGER_AREAS = options.ledgerAreas || DEFAULT_LEDGER_AREAS;
      this.LEDGER_ACTION_CODE_LEDGERS = options.ledgerActionCodeLedgers || DEFAULT_LEDGER_ACTION_CODE_LEDGERS;
    },

    /**
     * Process Excel file for asset changes with Ledger_Area format support
     */
    processExcelFile: function (file) {
      return new Promise((resolve, reject) => {
        if (!this._validateFileType(file)) {
          reject(new Error("Invalid file type. Please upload an Excel file."));
          return;
        }

        const reader = new FileReader();
        reader.onerror = (error) => {
          console.error("FileReader error:", error);
          MessageBox.error("Error reading file. Please try again.");
          reject(error);
        };

        reader.onload = (event) => {
          try {
            if (!window.XLSX) {
              throw new Error("XLSX library not loaded. Please ensure it's included in your project.");
            }

            const workbook = window.XLSX.read(event.target.result, {
              type: "array",
              cellDates: true,
              cellText: false,
              defval: ""
            });

            const sheetName = this._findChangeDataSheet(workbook);
            const parsedData = this._parseChangeSheet(workbook, sheetName);
            const validationResult = this._validateChangeEntries(parsedData);

            this._updateModels(validationResult);
            this._showProcessingResults(validationResult, file.name);
            this._clearFileUploader();

            resolve(validationResult);
          } catch (error) {
            console.error("Excel Processing Error", error);
            MessageBox.error(`Error processing file: ${file.name}`, {
              details: error.message
            });
            reject(error);
          }
        };

        reader.readAsArrayBuffer(file);
      });
    },

    /**
     * Parse Excel sheet data with enhanced Ledger_Area format support
     */
    _parseChangeSheet: function (workbook, sheetName) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        throw new Error(`Sheet "${sheetName}" not found`);
      }

      const sheetData = window.XLSX.utils.sheet_to_json(sheet, {
        raw: false,
        defval: null,
        dateNF: "yyyy-mm-dd",
        blankrows: false
      });

      if (sheetData.length === 0) {
        throw new Error(`No data found in sheet "${sheetName}"`);
      }

      console.log(`[ExcelProcessor] Processing ${sheetData.length} rows from sheet: ${sheetName}`);

      // Filter valid rows
      const filteredData = sheetData.filter((row, index) => {
        const sequenceNumber = row.SequenceNumber || row['Seq ID'] || row['Sequence Number'];
        const companyCode = row.CompanyCode || row['Company Code'];
        const fixedAsset = row.FixedAsset || row['Fixed Asset'] || row['Asset Number'];

        const hasBasicData = sequenceNumber && (companyCode || fixedAsset);
        if (!hasBasicData) {
          console.debug(`[ExcelProcessor] Skipping row ${index + 1}: insufficient basic data`);
          return false;
        }
        return true;
      });

      if (filteredData.length === 0) {
        throw new Error("No valid change data found. Please ensure your file contains proper asset change entries.");
      }

      // Transform data using DataTransformer with enhanced Ledger_Area support
      return filteredData.map((row, index) => {
        try {
          const changeStructure = this._dataTransformer.convertFlatToChangeStructure(row);

          if (!changeStructure.CompanyCode || !changeStructure.FixedAsset) {
            throw new Error(`Row ${index + 1}: Missing required fields CompanyCode or FixedAsset`);
          }

          if (!changeStructure.changes || Object.keys(changeStructure.changes).length === 0) {
            throw new Error(`Row ${index + 1}: No changes specified`);
          }

          return {
            SequenceNumber: changeStructure.SequenceNumber || (index + 1),
            CompanyCode: changeStructure.CompanyCode,
            FixedAsset: changeStructure.FixedAsset,
            ...changeStructure,
            // v3.0: preserve raw row so validators can access all columns
            _originalExcelData: row,
            Status: "Pending",
            _originalIndex: index,
            _sheetName: sheetName
          };

        } catch (error) {
          console.error(`Error processing row ${index + 1}:`, error);
          return {
            SequenceNumber: row.SequenceNumber || row['Seq ID'] || (index + 1),
            CompanyCode: row.CompanyCode || row['Company Code'] || "N/A",
            FixedAsset: row.FixedAsset || row['Fixed Asset'] || "N/A",
            // v3.0: preserve raw row even for error entries
            _originalExcelData: row,
            Status: "Invalid",
            ValidationErrors: [`Row processing error: ${error.message}`],
            _originalIndex: index,
            _rawData: row,
            _sheetName: sheetName
          };
        }
      });
    },

    /**
     * ✅ ENHANCED: Validate change entries with ValidityStartDate requirements
     */
    _validateChangeEntries: function (entries) {
      console.log("[ExcelProcessor] Validating", entries.length, "change entries");

      let validCount = 0;
      let invalidCount = 0;
      let warningCount = 0;

      const processedEntries = entries.map(entry => {
        if (entry.Status === "Invalid") {
          invalidCount++;
          return entry;
        }

        try {
          // Basic validation
          if (!entry.CompanyCode || !entry.FixedAsset) {
            entry.Status = "Invalid";
            entry.ValidationErrors = ["Missing required fields: CompanyCode or FixedAsset"];
            invalidCount++;
            return entry;
          }

          // Calculate changes summary using enhanced method with Ledger_Area support
          entry.ChangesSummary = this._dataTransformer._calculateChangesSummary(entry.changes || {});

          // ✅ NEW: Validate ValidityStartDate requirements
          const validityDateErrors = this._validateValidityStartDateRequirements(entry);

          // v3.0 NEW: Validate boolean-only fields (NegativeAmountIsAllowed, IsShutDown)
          const booleanFieldErrors = this._validateBooleanOnlyFields(entry);

          // Use DataTransformer validation
          let validationErrors = [];
          if (this._dataTransformer && typeof this._dataTransformer.validateAssetChangeData === 'function') {
            validationErrors = this._dataTransformer.validateAssetChangeData(entry);
          } else if (!entry.changes || Object.keys(entry.changes).length === 0) {
            validationErrors.push("No valid changes specified for this asset entry");
          }

          // ✅ NEW: Combine all validation errors
          const allErrors = [...validationErrors, ...validityDateErrors, ...booleanFieldErrors];

          // ✅ NEW: Separate critical errors from warnings
          const criticalErrors = allErrors.filter(err =>
            err.severity === 'Error' ||
            err.validationType === 'MANDATORY_FIELD' ||
            err.validationType === 'FIELD_FORMAT' ||
            err.validationType === 'BOOLEAN_FIELD_ERROR' || // v3.0 NEW
            typeof err === 'string' // plain strings from DataTransformer are critical
          );
          const warnings = allErrors.filter(err =>
            err.severity === 'Warning'
          );

          if (criticalErrors.length === 0) {
            if (warnings.length > 0) {
              entry.Status = "Valid"; // Still valid but has warnings
              entry.ValidationWarnings = warnings;
              warningCount++;
            } else {
              entry.Status = "Valid";
            }
            validCount++;
          } else {
            entry.Status = "Invalid";
            entry.ValidationErrors = criticalErrors;
            if (warnings.length > 0) {
              entry.ValidationWarnings = warnings;
            }
            invalidCount++;
          }

        } catch (error) {
          console.error("Validation error for entry", entry.SequenceNumber, error);
          entry.Status = "Invalid";
          entry.ValidationErrors = [`Internal validation error: ${error.message}`];
          invalidCount++;
        }

        return entry;
      });

      return {
        entries: processedEntries,
        totalCount: entries.length,
        validCount: validCount,
        invalidCount: invalidCount,
        warningCount: warningCount,
        isValid: invalidCount === 0
      };
    },

    /**
     * ✅ NEW: Validate ValidityStartDate requirements based on SAP documentation.
     * v3.0: Now uses this.LEDGER_AREAS (dynamic) and also validates
     * TimeBasedValuation_ActionCode independently from Valuation_ActionCode.
     *
     * SAP Documentation: ValidityStartDate has Cardinality 1..1 (MANDATORY for CREATE)
     * @private
     */
    _validateValidityStartDateRequirements: function (entry) {
      const errors = [];
      const sequenceId = entry.SequenceNumber || 'N/A';
      const asset = entry.FixedAsset || 'Unknown';
      const originalData = entry._originalExcelData || {};

      // v3.0: use configured LEDGER_AREAS instead of hardcoded array
      this.LEDGER_AREAS.forEach(ledgerArea => {
        const valuationAction  = entry.actionCodes?.valuations?.[ledgerArea];
        const timeBasedAction  = entry.actionCodes?.timeBasedValuations?.[ledgerArea]; // v3.0 NEW

        // Only validate if there is a valuation or time-based action
        if (!valuationAction && !timeBasedAction) return;

        // Check for ledger-area-specific ValidityStartDate
        const explicitValidityDate =
          originalData[`ValidityStartDate_${ledgerArea}_Value`] ||
          originalData[`ValidityStartDate_${ledgerArea}`];

        // Check for general fallback dates
        const fallbackDates =
          originalData['ValidityStartDate_Value'] ||
          originalData['ValidityStartDate'] ||
          originalData['IN_AssetPutToUseDate_Value'] ||
          originalData['IN_AssetPutToUseDate'] ||
          originalData[`DepreciationStartDate_${ledgerArea}_Value`] ||
          originalData[`DepreciationStartDate_${ledgerArea}`] ||
          originalData['DepreciationStartDate_Value'] ||
          originalData['DepreciationStartDate'];

        const effectiveDate = explicitValidityDate || fallbackDates;

        // ── Helpers ──────────────────────────────────────────────────────────

        const pushMissingDateError = (actionLabel) => {
          errors.push({
            field: `ValidityStartDate_${ledgerArea}`,
            message:
              `REQUIRED: ValidityStartDate is mandatory when ${actionLabel} for ${ledgerArea}. ` +
              `Add column "ValidityStartDate_${ledgerArea}_Value" with a NEW date (e.g., 2026-01-20) ` +
              `or provide a general "ValidityStartDate_Value" column.`,
            severity: 'Error',
            validationType: 'MANDATORY_FIELD',
            sequenceId,
            asset
          });
        };

        const pushFormatError = (dateValue, context) => {
          errors.push({
            field: `ValidityStartDate_${ledgerArea}`,
            message:
              `Invalid date format for ValidityStartDate in ${ledgerArea}${context ? ` (${context})` : ''}. ` +
              `Use YYYY-MM-DD format (e.g., 2026-01-20).`,
            severity: 'Error',
            validationType: 'FIELD_FORMAT',
            sequenceId,
            asset
          });
        };

        const pushChangeWarning = (dateValue, actionOwner) => {
          errors.push({
            field: `ValidityStartDate_${ledgerArea}`,
            message:
              `⚠️ IMPORTANT: When using CHANGE action (02) for ${ledgerArea}` +
              (actionOwner ? ` [${actionOwner}]` : '') +
              `, ValidityStartDate (${dateValue}) must EXACTLY match an existing interval in SAP. ` +
              `If this date doesn't exist, SAP will return error: ` +
              `"Dep. area ${ledgerArea.split('_')[1]} of value base time interval starting with ${dateValue} doesn't exist". ` +
              `To avoid this error, either: (1) Leave ValidityStartDate blank and SAP will use the existing date, ` +
              `OR (2) Ensure this date matches exactly.`,
            severity: 'Warning',
            validationType: 'VALIDITY_DATE_WARNING',
            sequenceId,
            asset
          });
        };

        // ── Valuation-level rules (same as v2.6) ─────────────────────────────

        // ✅ VALIDATION RULE 1: CREATE (01) - ValidityStartDate is MANDATORY
        if (valuationAction === "01") {
          if (!effectiveDate) {
            pushMissingDateError("creating new depreciation");
          } else if (!this._isValidDateFormat(effectiveDate)) {
            pushFormatError(effectiveDate);
          }
        }

        // ✅ VALIDATION RULE 2: CHANGE (02) - ValidityStartDate should match existing SAP date
        if (valuationAction === "02" && explicitValidityDate) {
          if (!this._isValidDateFormat(explicitValidityDate)) {
            pushFormatError(explicitValidityDate);
          } else {
            pushChangeWarning(explicitValidityDate, "Valuation");
          }
          // Note: If no ValidityStartDate provided for CHANGE, that's OK - SAP will use existing date
        }

        // ── v3.0 NEW: TimeBasedValuation-level rules ──────────────────────────
        // TimeBasedValuation_ActionCode is INDEPENDENT from Valuation_ActionCode.
        // Example: Valuation=02 (CHANGE) + TimeBasedValuation=01 (CREATE) → date still required.
        // Only validate if TBV action is present and differs from Valuation action
        // (same action is already covered above).
        if (timeBasedAction && timeBasedAction !== valuationAction) {
          if (timeBasedAction === "01") {
            if (!effectiveDate) {
              pushMissingDateError("creating new time-based valuation");
            } else if (!this._isValidDateFormat(effectiveDate)) {
              pushFormatError(effectiveDate, "TimeBasedValuation CREATE");
            }
          } else if (timeBasedAction === "02" && explicitValidityDate) {
            if (!this._isValidDateFormat(explicitValidityDate)) {
              pushFormatError(explicitValidityDate, "TimeBasedValuation CHANGE");
            } else {
              pushChangeWarning(explicitValidityDate, "TimeBasedValuation");
            }
          }
        }
      });

      return errors;
    },

    /**
     * v3.0 NEW: Validate boolean-only fields (NegativeAmountIsAllowed, IsShutDown).
     * These have no _Change counterpart — value must be "true" or "false" if provided.
     * @private
     */
    _validateBooleanOnlyFields: function (entry) {
      const errors = [];
      const sequenceId = entry.SequenceNumber || 'N/A';
      const asset = entry.FixedAsset || 'Unknown';
      const originalData = entry._originalExcelData || {};
      const validValues = ["true", "false", "TRUE", "FALSE", "True", "False", "1", "0", ""];

      BOOLEAN_ONLY_FIELDS.forEach(fieldBase => {
        this.LEDGER_AREAS.forEach(ledgerArea => {
          const colKey = `${fieldBase}_${ledgerArea}`;
          const rawVal = originalData[colKey];

          if (rawVal === null || rawVal === undefined || rawVal === "") return; // blank is fine

          if (!validValues.includes(String(rawVal).trim())) {
            errors.push({
              field: colKey,
              message:
                `Invalid value "${rawVal}" for boolean field ${colKey}. ` +
                `Allowed values: true / false. No _Change column is needed for this field.`,
              severity: 'Error',
              validationType: 'BOOLEAN_FIELD_ERROR',
              ledgerArea,
              sequenceId,
              asset
            });
          }
        });
      });

      return errors;
    },

    /**
     * ✅ NEW: Validate date format (YYYY-MM-DD)
     * @private
     */
    _isValidDateFormat: function (dateValue) {
      if (!dateValue) return false;

      // Check YYYY-MM-DD format
      const datePattern = /^\d{4}-\d{2}-\d{2}$/;
      if (!datePattern.test(dateValue)) {
        return false;
      }

      // Validate it's a real date
      const date = new Date(dateValue);
      return !isNaN(date.getTime());
    },

    /**
     * Update models with correct naming and enhanced Ledger_Area support
     */
    _updateModels: function (validationResult) {
      const controller = this.oController;
      const view = controller.getView();

      console.log("[ExcelProcessor] Updating models with", validationResult.entries.length, "entries");

      // Use correct model name that matches view binding
      let assetMasterModel = view.getModel("assetMasterEntries");
      if (!assetMasterModel) {
        console.warn("[ExcelProcessor] Creating missing assetMasterEntries model");
        assetMasterModel = new JSONModel();
        view.setModel(assetMasterModel, "assetMasterEntries");
      }

      // Set data correctly
      assetMasterModel.setProperty("/entries", validationResult.entries);
      assetMasterModel.setProperty("/validationStatus", validationResult.isValid ? "Valid" : "Invalid");
      assetMasterModel.setProperty("/filteredCount", validationResult.totalCount);
      assetMasterModel.refresh(true);

      console.log("[ExcelProcessor] assetMasterEntries model updated with", validationResult.entries.length, "entries");

      // Update upload summary
      let uploadSummaryModel = view.getModel("uploadSummary");
      if (!uploadSummaryModel) {
        console.warn("[ExcelProcessor] Creating missing uploadSummary model");
        uploadSummaryModel = new JSONModel();
        view.setModel(uploadSummaryModel, "uploadSummary");
      }

      uploadSummaryModel.setProperty("/TotalEntries", validationResult.totalCount);
      uploadSummaryModel.setProperty("/SuccessfulEntries", validationResult.validCount);
      uploadSummaryModel.setProperty("/FailedEntries", validationResult.invalidCount);
      uploadSummaryModel.setProperty("/WarningCount", validationResult.warningCount || 0);
      uploadSummaryModel.setProperty("/LastUploadDate", new Date());
      uploadSummaryModel.setProperty("/ProcessingComplete", true);
      uploadSummaryModel.setProperty("/IsSubmitEnabled", validationResult.validCount > 0);
      uploadSummaryModel.setProperty("/HasBeenSubmitted", false);
      uploadSummaryModel.refresh(true);

      // Update change summary model with Ledger_Area support
      this._updateChangeSummaryModelWithLedgerArea(validationResult.entries);

      // Force table refresh
      setTimeout(() => {
        const table = controller.byId("assetChangeTable");
        if (table) {
          const binding = table.getBinding("items");
          if (binding) {
            binding.refresh();
            console.log("[ExcelProcessor] Table binding refreshed");
          }
        }
      }, 100);
    },

    /**
     * Update change summary model with enhanced Ledger_Area format support.
     * v3.0: ledgerAreaSummary is now dynamic (not hardcoded to 4 areas),
     * timeBasedValuation action code counts added, booleanFieldIssues tracked.
     */
    _updateChangeSummaryModelWithLedgerArea: function (entries) {
      const controller = this.oController;
      let changeSummaryModel = controller.getView().getModel("changeSummary");

      if (!changeSummaryModel) {
        changeSummaryModel = new JSONModel();
        controller.getView().setModel(changeSummaryModel, "changeSummary");
      }

      // v3.0: build ledgerAreaSummary dynamically from configured areas
      const ledgerAreaSummary = {};
      this.LEDGER_AREAS.forEach(la => {
        ledgerAreaSummary[la] = { ledger: 0, valuation: 0, timeBasedValuation: 0 }; // v3.0: +timeBasedValuation
      });

      const summaryData = {
        totalAssets: entries.length,
        assetsWithChanges: 0,
        fieldChangeSummary: {},
        blankValueAssignments: [],
        validityDateWarnings: [], // ✅ NEW: Track ValidityStartDate warnings
        booleanFieldIssues: [],   // v3.0 NEW: Track boolean field errors
        actionCodeSummary: {
          accountAssignment: { create: 0, change: 0 },
          ledger:            { create: 0, change: 0, delete: 0 },
          valuation:         { create: 0, change: 0, delete: 0 },
          timeBasedValuation:{ create: 0, change: 0, delete: 0 } // v3.0 NEW
        },
        // Enhanced tracking for Ledger_Area format (now dynamic)
        ledgerAreaSummary
      };

      entries.forEach(entry => {
        let hasChanges = false;

        if (entry.changes && Object.keys(entry.changes).length > 0) {
          Object.entries(entry.changes).forEach(([fieldName, changeInfo]) => {
            if (changeInfo && changeInfo.shouldChange) {
              hasChanges = true;

              if (!summaryData.fieldChangeSummary[fieldName]) {
                summaryData.fieldChangeSummary[fieldName] = 0;
              }
              summaryData.fieldChangeSummary[fieldName]++;

              if (changeInfo.isBlankValue) {
                summaryData.blankValueAssignments.push({
                  asset: entry.FixedAsset,
                  sequenceNumber: entry.SequenceNumber,
                  field: fieldName,
                  companyCode: entry.CompanyCode
                });
              }
            }
          });
        }

        // ✅ NEW: Track ValidityStartDate warnings
        if (entry.ValidationWarnings && Array.isArray(entry.ValidationWarnings)) {
          entry.ValidationWarnings.forEach(warning => {
            if (warning.validationType === 'VALIDITY_DATE_WARNING') {
              summaryData.validityDateWarnings.push({
                asset: entry.FixedAsset,
                sequenceNumber: entry.SequenceNumber,
                field: warning.field,
                message: warning.message,
                companyCode: entry.CompanyCode
              });
            }
          });
        }

        // v3.0 NEW: Track boolean field errors
        if (entry.ValidationErrors && Array.isArray(entry.ValidationErrors)) {
          entry.ValidationErrors.forEach(err => {
            if (err.validationType === 'BOOLEAN_FIELD_ERROR') {
              summaryData.booleanFieldIssues.push({
                asset: entry.FixedAsset,
                sequenceNumber: entry.SequenceNumber,
                field: err.field,
                message: err.message,
                companyCode: entry.CompanyCode
              });
            }
          });
        }

        // Count action codes with enhanced Ledger_Area tracking
        if (entry.actionCodes) {
          // Account Assignment
          if (entry.actionCodes.accountAssignment) {
            const action = entry.actionCodes.accountAssignment;
            if (action === "01") summaryData.actionCodeSummary.accountAssignment.create++;
            else if (action === "02") summaryData.actionCodeSummary.accountAssignment.change++;
          }

          // Ledgers with Ledger_Area format
          Object.entries(entry.actionCodes.ledgers || {}).forEach(([ledgerArea, action]) => {
            if (action === "01") summaryData.actionCodeSummary.ledger.create++;
            else if (action === "02") summaryData.actionCodeSummary.ledger.change++;
            else if (action === "03") summaryData.actionCodeSummary.ledger.delete++;

            // Track specific Ledger_Area activity
            if (summaryData.ledgerAreaSummary[ledgerArea]) {
              summaryData.ledgerAreaSummary[ledgerArea].ledger++;
            }
          });

          // Valuations with Ledger_Area format
          Object.entries(entry.actionCodes.valuations || {}).forEach(([ledgerArea, action]) => {
            if (action === "01") summaryData.actionCodeSummary.valuation.create++;
            else if (action === "02") summaryData.actionCodeSummary.valuation.change++;
            else if (action === "03") summaryData.actionCodeSummary.valuation.delete++;

            // Track specific Ledger_Area activity
            if (summaryData.ledgerAreaSummary[ledgerArea]) {
              summaryData.ledgerAreaSummary[ledgerArea].valuation++;
            }
          });

          // v3.0 NEW: TimeBasedValuation with Ledger_Area format (independent action code)
          Object.entries(entry.actionCodes.timeBasedValuations || {}).forEach(([ledgerArea, action]) => {
            if (action === "01") summaryData.actionCodeSummary.timeBasedValuation.create++;
            else if (action === "02") summaryData.actionCodeSummary.timeBasedValuation.change++;
            else if (action === "03") summaryData.actionCodeSummary.timeBasedValuation.delete++;

            if (summaryData.ledgerAreaSummary[ledgerArea]) {
              summaryData.ledgerAreaSummary[ledgerArea].timeBasedValuation++;
            }
          });
        }

        if (hasChanges) {
          summaryData.assetsWithChanges++;
        }
      });

      changeSummaryModel.setData(summaryData);
      changeSummaryModel.refresh(true);
    },

    /**
     * ✅ ENHANCED: Show processing results with ValidityStartDate warnings.
     * v3.0: also surfaces TimeBasedValuation CREATE counts and boolean field errors.
     */
    _showProcessingResults: function (validationResult, fileName) {
      let totalChanges = 0;
      let blankAssignments = 0;
      let createCount = 0;
      let changeCount = 0;
      let deleteCount = 0;
      let validityDateWarningCount = 0;
      let tbvCreateCount = 0;           // v3.0 NEW
      let booleanFieldErrorCount = 0;   // v3.0 NEW
      let ledgerAreaActivity = {};

      validationResult.entries.forEach(entry => {
        if (entry.ChangesSummary) {
          totalChanges += entry.ChangesSummary.totalChanges || 0;
          blankAssignments += entry.ChangesSummary.blankAssignments || 0;
          createCount += entry.ChangesSummary.createCount || 0;
          changeCount += entry.ChangesSummary.changeCount || 0;
          deleteCount += entry.ChangesSummary.deleteCount || 0;

          // Track Ledger_Area specific activity
          if (entry.ChangesSummary.ledgerActions) {
            Object.keys(entry.ChangesSummary.ledgerActions).forEach(ledgerArea => {
              if (!ledgerAreaActivity[ledgerArea]) ledgerAreaActivity[ledgerArea] = 0;
              ledgerAreaActivity[ledgerArea]++;
            });
          }

          if (entry.ChangesSummary.valuationActions) {
            Object.keys(entry.ChangesSummary.valuationActions).forEach(ledgerArea => {
              if (!ledgerAreaActivity[ledgerArea]) ledgerAreaActivity[ledgerArea] = 0;
              ledgerAreaActivity[ledgerArea]++;
            });
          }

          // v3.0 NEW: TimeBasedValuation area activity
          if (entry.ChangesSummary.timeBasedValuationActions) {
            Object.keys(entry.ChangesSummary.timeBasedValuationActions).forEach(ledgerArea => {
              if (!ledgerAreaActivity[ledgerArea]) ledgerAreaActivity[ledgerArea] = 0;
              ledgerAreaActivity[ledgerArea]++;
            });
          }
        }

        // ✅ NEW: Count ValidityStartDate warnings
        if (entry.ValidationWarnings && Array.isArray(entry.ValidationWarnings)) {
          validityDateWarningCount += entry.ValidationWarnings.filter(w =>
            w.validationType === 'VALIDITY_DATE_WARNING'
          ).length;
        }

        // v3.0 NEW: Count boolean field errors and TBV creates
        if (entry.ValidationErrors && Array.isArray(entry.ValidationErrors)) {
          booleanFieldErrorCount += entry.ValidationErrors.filter(e =>
            e.validationType === 'BOOLEAN_FIELD_ERROR'
          ).length;
        }

        Object.values(entry.actionCodes?.timeBasedValuations || {}).forEach(action => {
          if (action === "01") tbvCreateCount++;
        });
      });

      let message = `File processed: ${validationResult.totalCount} assets found ` +
        `(${validationResult.validCount} valid, ${validationResult.invalidCount} invalid`;

      if (validationResult.warningCount > 0) {
        message += `, ${validationResult.warningCount} with warnings`;
      }
      message += `)`;

      if (totalChanges > 0) {
        message += `\n${totalChanges} changes detected`;

        const actions = [];
        if (createCount > 0) actions.push(`${createCount} create`);
        if (changeCount > 0) actions.push(`${changeCount} change`);
        if (deleteCount > 0) actions.push(`${deleteCount} delete`);
        if (tbvCreateCount > 0) actions.push(`${tbvCreateCount} TBV-create`); // v3.0

        if (actions.length > 0) {
          message += ` (${actions.join(", ")})`;
        }

        // Add Ledger_Area specific information
        const activeLedgerAreas = Object.keys(ledgerAreaActivity);
        if (activeLedgerAreas.length > 0) {
          message += `\nLedger areas affected: ${activeLedgerAreas.join(", ")}`;
        }

        if (blankAssignments > 0) {
          message += `\n${blankAssignments} blank value assignments`;
        }
      }

      MessageToast.show(message, { duration: 8000 });

      // Show validation errors if any
      if (!validationResult.isValid && this._uiManager && typeof this._uiManager.handleValidationErrors === 'function') {
        const invalidEntries = validationResult.entries.filter(e => e.Status === 'Invalid');
        const allErrors = invalidEntries.flatMap(entry =>
          (entry.ValidationErrors || []).map(error => {
            if (typeof error === 'string') {
              return {
                sequenceId: entry.SequenceNumber,
                asset: entry.FixedAsset || 'N/A',
                companyCode: entry.CompanyCode || 'N/A',
                message: error,
                field: 'General',
                severity: 'Error'
              };
            }
            return {
              sequenceId: error.sequenceId || entry.SequenceNumber,
              asset: error.asset || entry.FixedAsset || 'N/A',
              companyCode: error.companyCode || entry.CompanyCode || 'N/A',
              message: error.message,
              field: error.field || 'General',
              severity: error.severity || 'Error'
            };
          })
        );

        this._uiManager.handleValidationErrors(allErrors);
      }

      // ✅ NEW: Show ValidityStartDate warnings
      if (validityDateWarningCount > 0) {
        MessageBox.warning(
          `${validityDateWarningCount} ValidityStartDate warning(s) detected. ` +
          `For CHANGE actions (02), ensure the dates match existing intervals in SAP. ` +
          `If unsure, leave ValidityStartDate blank to use SAP's existing dates.\n\n` +
          `This also applies when TimeBasedValuation_ActionCode differs from Valuation_ActionCode ` +
          `(e.g., Valuation=02 + TimeBasedValuation=01 still requires a valid date).`,
          {
            title: "ValidityStartDate Warning",
            details: "Click on asset entries to view detailed warnings."
          }
        );
      }

      // v3.0 NEW: Show boolean field errors
      if (booleanFieldErrorCount > 0) {
        MessageBox.error(
          `${booleanFieldErrorCount} boolean field error(s) detected.\n` +
          `Fields NegativeAmountIsAllowed and IsShutDown only accept true or false.\n` +
          `These fields have no _Change column — set the value directly.`,
          { title: "Boolean Field Validation Error" }
        );
      }

      if (blankAssignments > 0) {
        MessageBox.warning(
          `${blankAssignments} field(s) will be set to blank values. Please review carefully before submission.`,
          { title: "Blank Value Assignment Warning" }
        );
      }

      if (deleteCount > 0) {
        MessageBox.warning(
          `${deleteCount} deletion action(s) detected. This will permanently remove data. Please review carefully before submission.`,
          { title: "Deletion Action Warning" }
        );
      }

      // Show Ledger_Area specific warnings
      const activeLedgerAreas = Object.keys(ledgerAreaActivity);
      const criticalLedgerAreas = ['0L_01', '2L_32']; // Primary book and consolidation
      const criticalAffected = activeLedgerAreas.filter(area => criticalLedgerAreas.includes(area));
      if (criticalAffected.length > 0) {
        MessageBox.information(
          `Changes detected in critical ledger areas: ${criticalAffected.join(", ")}. Please ensure these changes are reviewed by appropriate personnel.`,
          { title: "Critical Ledger Area Changes" }
        );
      }
    },

    /**
     * Export change summary with enhanced Ledger_Area format support.
     * v3.0: TimeBasedValuation rows in Action Codes sheet, HasLedgerActionCode column,
     * boolean field issues sheet, TBVCreateActions in Ledger Area Summary.
     */
    exportChangeSummary: function (entries) {
      if (!window.XLSX) {
        throw new Error("XLSX library not loaded");
      }

      const wb = window.XLSX.utils.book_new();
      const summaryData = [];
      const detailData = [];
      const actionSummaryData = [];
      const ledgerAreaSummaryData = [];
      const validityDateWarningsData = []; // ✅ NEW: ValidityStartDate warnings sheet
      const booleanFieldIssuesData = [];   // v3.0 NEW: Boolean field issues sheet

      entries.forEach(entry => {
        let assetHasChanges = false;

        if (entry.changes && Object.keys(entry.changes).length > 0) {
          Object.entries(entry.changes).forEach(([fieldName, changeInfo]) => {
            if (changeInfo && changeInfo.shouldChange) {
              assetHasChanges = true;

              // Determine action
              let actionCode = "02";
              let actionName = "Change";
              let ledgerArea = "";

              if (changeInfo.ledgerActionCode) {
                actionCode = changeInfo.ledgerActionCode;
                actionName = this._getActionName(actionCode);
                ledgerArea = changeInfo.ledgerAreaCode || "";
              } else if (changeInfo.valuationActionCode) {
                actionCode = changeInfo.valuationActionCode;
                actionName = this._getActionName(actionCode);
                ledgerArea = changeInfo.ledgerAreaCode || "";
              } else if (changeInfo.timeBasedValuationActionCode) { // v3.0 NEW
                actionCode = changeInfo.timeBasedValuationActionCode;
                actionName = this._getActionName(actionCode);
                ledgerArea = changeInfo.ledgerAreaCode || "";
              }

              summaryData.push({
                SequenceNumber: entry.SequenceNumber,
                CompanyCode: entry.CompanyCode,
                FixedAsset: entry.FixedAsset,
                Status: entry.Status,
                FieldToChange: fieldName,
                NewValue: changeInfo.isBlankValue ? "<BLANK>" : changeInfo.newValue,
                IsBlankAssignment: changeInfo.isBlankValue ? "Yes" : "No",
                Action: actionName,
                ActionCode: actionCode,
                LedgerArea: ledgerArea,
                LedgerCode: changeInfo.ledgerCode || "",
                AreaCode: changeInfo.areaCode || ""
              });
            }
          });
        }

        // ✅ NEW: Collect ValidityStartDate warnings
        if (entry.ValidationWarnings && Array.isArray(entry.ValidationWarnings)) {
          entry.ValidationWarnings.forEach(warning => {
            if (warning.validationType === 'VALIDITY_DATE_WARNING') {
              validityDateWarningsData.push({
                SequenceNumber: entry.SequenceNumber,
                CompanyCode: entry.CompanyCode,
                FixedAsset: entry.FixedAsset,
                Field: warning.field,
                Warning: warning.message,
                Severity: warning.severity,
                Recommendation: "Leave ValidityStartDate blank to use existing SAP date, OR ensure exact match"
              });
            }
          });
        }

        // v3.0 NEW: Collect boolean field issues
        if (entry.ValidationErrors && Array.isArray(entry.ValidationErrors)) {
          entry.ValidationErrors.forEach(err => {
            if (err.validationType === 'BOOLEAN_FIELD_ERROR') {
              booleanFieldIssuesData.push({
                SequenceNumber: entry.SequenceNumber,
                CompanyCode: entry.CompanyCode,
                FixedAsset: entry.FixedAsset,
                Field: err.field,
                LedgerArea: err.ledgerArea || "",
                Error: err.message,
                Fix: "Change value to true or false. No _Change column needed."
              });
            }
          });
        }

        // Add action codes summary with Ledger_Area format
        if (entry.actionCodes) {
          if (entry.actionCodes.accountAssignment) {
            actionSummaryData.push({
              SequenceNumber: entry.SequenceNumber,
              CompanyCode: entry.CompanyCode,
              FixedAsset: entry.FixedAsset,
              StructureLevel: "Account Assignment",
              ActionCode: entry.actionCodes.accountAssignment,
              Action: this._getActionName(entry.actionCodes.accountAssignment),
              LedgerArea: "N/A",
              HasLedgerActionCode: "N/A",   // v3.0 NEW
              IsTimeBasedValuation: "NO"    // v3.0 NEW
            });
          }

          Object.entries(entry.actionCodes.ledgers || {}).forEach(([ledgerArea, code]) => {
            actionSummaryData.push({
              SequenceNumber: entry.SequenceNumber,
              CompanyCode: entry.CompanyCode,
              FixedAsset: entry.FixedAsset,
              StructureLevel: "Ledger",
              ActionCode: code,
              Action: this._getActionName(code),
              LedgerArea: ledgerArea,
              HasLedgerActionCode: "YES",   // v3.0: ledger rows always have it
              IsTimeBasedValuation: "NO"
            });
          });

          Object.entries(entry.actionCodes.valuations || {}).forEach(([ledgerArea, code]) => {
            const ledger = ledgerArea.split("_")[0];
            actionSummaryData.push({
              SequenceNumber: entry.SequenceNumber,
              CompanyCode: entry.CompanyCode,
              FixedAsset: entry.FixedAsset,
              StructureLevel: "Valuation",
              ActionCode: code,
              Action: this._getActionName(code),
              LedgerArea: ledgerArea,
              // v3.0: flag whether this area's ledger has Ledger_ActionCode
              HasLedgerActionCode: this.LEDGER_ACTION_CODE_LEDGERS.includes(ledger) ? "YES" : "NO",
              IsTimeBasedValuation: "NO"
            });
          });

          // v3.0 NEW: TimeBasedValuation rows in Action Codes sheet
          Object.entries(entry.actionCodes.timeBasedValuations || {}).forEach(([ledgerArea, code]) => {
            const ledger = ledgerArea.split("_")[0];
            actionSummaryData.push({
              SequenceNumber: entry.SequenceNumber,
              CompanyCode: entry.CompanyCode,
              FixedAsset: entry.FixedAsset,
              StructureLevel: "TimeBasedValuation",
              ActionCode: code,
              Action: this._getActionName(code),
              LedgerArea: ledgerArea,
              HasLedgerActionCode: this.LEDGER_ACTION_CODE_LEDGERS.includes(ledger) ? "YES" : "NO",
              IsTimeBasedValuation: "YES"
            });
          });
        }

        detailData.push({
          SequenceNumber: entry.SequenceNumber,
          CompanyCode: entry.CompanyCode,
          FixedAsset: entry.FixedAsset,
          Status: entry.Status,
          TotalChanges: entry.ChangesSummary ? entry.ChangesSummary.totalChanges : 0,
          CreateActions: entry.ChangesSummary ? entry.ChangesSummary.createCount : 0,
          ChangeActions: entry.ChangesSummary ? entry.ChangesSummary.changeCount : 0,
          DeleteActions: entry.ChangesSummary ? entry.ChangesSummary.deleteCount : 0,
          BlankAssignments: entry.ChangesSummary ? entry.ChangesSummary.blankAssignments : 0,
          ValidationErrors: (entry.ValidationErrors || []).map(e =>
            typeof e === 'string' ? e : (e.message || 'Unknown Error')
          ).join('; '),
          ValidationWarnings: (entry.ValidationWarnings || []).map(w =>
            typeof w === 'string' ? w : (w.message || 'Unknown Warning')
          ).join('; ')
        });
      });

      // Create Ledger_Area summary
      const ledgerAreaStats = {};
      actionSummaryData.forEach(action => {
        if (action.LedgerArea !== "N/A") {
          if (!ledgerAreaStats[action.LedgerArea]) {
            ledgerAreaStats[action.LedgerArea] = {
              LedgerArea: action.LedgerArea,
              Ledger: action.LedgerArea.split("_")[0],        // v3.0 NEW
              DeprecArea: action.LedgerArea.split("_")[1],    // v3.0 NEW
              HasLedgerActionCode: action.HasLedgerActionCode,// v3.0 NEW
              TotalActions: 0,
              CreateActions: 0,
              ChangeActions: 0,
              DeleteActions: 0,
              TBVCreateActions: 0, // v3.0 NEW
              AffectedAssets: new Set()
            };
          }
          const stat = ledgerAreaStats[action.LedgerArea];
          stat.TotalActions++;
          stat.AffectedAssets.add(action.FixedAsset);

          if (action.ActionCode === "01") {
            action.IsTimeBasedValuation === "YES"
              ? stat.TBVCreateActions++
              : stat.CreateActions++;
          } else if (action.ActionCode === "02") stat.ChangeActions++;
          else if (action.ActionCode === "03") stat.DeleteActions++;
        }
      });

      Object.values(ledgerAreaStats).forEach(stat => {
        ledgerAreaSummaryData.push({
          LedgerArea: stat.LedgerArea,
          Ledger: stat.Ledger,                      // v3.0 NEW
          DeprecArea: stat.DeprecArea,              // v3.0 NEW
          HasLedgerActionCode: stat.HasLedgerActionCode, // v3.0 NEW
          TotalActions: stat.TotalActions,
          CreateActions: stat.CreateActions,
          ChangeActions: stat.ChangeActions,
          DeleteActions: stat.DeleteActions,
          TBVCreateActions: stat.TBVCreateActions,  // v3.0 NEW
          AffectedAssets: stat.AffectedAssets.size
        });
      });

      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(summaryData), "Change Summary");
      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(detailData), "Asset Details");
      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(actionSummaryData), "Action Codes");
      window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(ledgerAreaSummaryData), "Ledger Area Summary");

      // ✅ NEW: Add ValidityStartDate warnings sheet
      if (validityDateWarningsData.length > 0) {
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(validityDateWarningsData), "ValidityDate Warnings");
      }

      // v3.0 NEW: Add boolean field issues sheet
      if (booleanFieldIssuesData.length > 0) {
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(booleanFieldIssuesData), "Boolean Field Issues");
      }

      const timestamp = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '');
      const filename = `Asset_Change_Summary_LedgerArea_${timestamp}.xlsx`;

      window.XLSX.writeFile(wb, filename);
      MessageToast.show(`Change summary exported to ${filename}`);
    },

    _getActionName: function (actionCode) {
      switch (actionCode) {
        case "01": return "Create";
        case "02": return "Change";
        case "03": return "Delete";
        default: return "Change";
      }
    },

    // === UTILITY METHODS ===

    _validateFileType: function (file) {
      const validMimeTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv'
      ];
      return validMimeTypes.includes(file.type) || /\.(xlsx|xls|csv)$/i.test(file.name);
    },

    _findChangeDataSheet: function (workbook) {
      const sheetNames = workbook.SheetNames;
      const searchPriority = ['Change_Data', 'change_data', 'ChangeData', 'Data', 'Template', 'Main', 'Assets'];

      for (const targetName of searchPriority) {
        const found = sheetNames.find(name => name.toLowerCase() === targetName.toLowerCase());
        if (found) {
          console.log(`[ExcelProcessor] Using sheet: ${found}`);
          return found;
        }
      }

      const changeRelatedSheet = sheetNames.find(name =>
        name.toLowerCase().includes('change') ||
        name.toLowerCase().includes('data') ||
        name.toLowerCase().includes('template') ||
        name.toLowerCase().includes('asset')
      );

      if (changeRelatedSheet) {
        console.log(`[ExcelProcessor] Using change-related sheet: ${changeRelatedSheet}`);
        return changeRelatedSheet;
      }

      console.log(`[ExcelProcessor] Using first sheet: ${sheetNames[0]}`);
      return sheetNames[0];
    },

    _clearFileUploader: function () {
      const fileUploader = this.oController.getView().byId("fileUploader");
      if (fileUploader && typeof fileUploader.clear === 'function') {
        fileUploader.clear();
      }
    }
  });
});