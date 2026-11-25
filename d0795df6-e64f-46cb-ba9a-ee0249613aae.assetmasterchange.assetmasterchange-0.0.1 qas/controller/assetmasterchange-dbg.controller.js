sap.ui.define([
        "./BaseFilterController",
        "sap/ui/model/json/JSONModel",
        "sap/m/MessageBox",
        "sap/m/MessageToast",
        "assetmasterchange/model/ModelManager",
        "assetmasterchange/service/ExcelProcessor",
        "assetmasterchange/service/UIManager",
        "assetmasterchange/utils/ErrorHandler",
        "assetmasterchange/service/ExportManager",
        "assetmasterchange/service/BatchProcessingManager",
        "assetmasterchange/utils/DataTransformer",
        "assetmasterchange/service/UserService",
        "assetmasterchange/service/ODataService"
], function (
        BaseFilterController,
        JSONModel,
        MessageBox,
        MessageToast,
        ModelManager,
        ExcelProcessor,
        UIManager,
        ErrorHandler,
        ExportManager,
        BatchProcessingManager,
        DataTransformer,
        UserService,
        ODataService
) {
        "use strict";

        return BaseFilterController.extend("assetmasterchange.controller.AssetMasterChange", {

                /**
                 * Controller initialization
                 */
                onInit: function () {
                        this.onInitFilters();
                        this._initializeModels();
                        this._initializeServices();
                        this._userService.getUserInfo();
                },

                /**
                 * Initialize application models
                 * @private
                 */
                _initializeModels: function () {
                        this._modelManager = new ModelManager(this);
                        // Upload summary model
                        const uploadSummaryModel = new JSONModel({
                                TotalEntries: 0,
                                SuccessfulEntries: 0,
                                FailedEntries: 0,
                                WarningCount: 0, // ✅ NEW: Track warnings
                                ValidationErrors: [],
                                ValidationWarnings: [], // ✅ NEW: Store warnings separately
                                IsSubmitEnabled: false,
                                HasBeenSubmitted: false,
                                ProcessingComplete: false,
                                isSubmitting: false
                        });
                        this.getView().setModel(uploadSummaryModel, "uploadSummary");

                        // Asset master entries model
                        const assetMasterModel = new JSONModel({
                                entries: [],
                                validationStatus: "",
                                filteredCount: 0
                        });
                        this.getView().setModel(assetMasterModel, "assetMasterEntries");
                },

                /**
                 * Initialize application services
                 * @private
                 */
                _initializeServices: function () {
                        this._errorHandler = new ErrorHandler();
                        this._dataTransformer = new DataTransformer();
                        this._userService = new UserService(this);

                        this._exportManager = new ExportManager({
                                errorHandler: this._errorHandler,
                                dataTransformer: this._dataTransformer
                        });
                        this._uiManager = new UIManager(this, this._errorHandler, this._exportManager);

                        const oModel = this.getOwnerComponent().getModel();
                        this._oDataService = new ODataService(oModel, this._dataTransformer, this._errorHandler);
                        this._oDataService.setDebugMode(true);
                        this._batchProcessingManager = new BatchProcessingManager({
                                controller: this,
                                uiManager: this._uiManager,
                                errorHandler: this._errorHandler,
                                dataTransformer: this._dataTransformer,
                                oDataService: this._oDataService,
                                ExportManager: this._exportManager
                        });
                        this._excelProcessor = new ExcelProcessor({
                                controller: this,
                                uiManager: this._uiManager,
                                dataTransformer: this._dataTransformer
                        });
                        console.log("Asset master change services initialized successfully");
                },

                /**
                 * Handle file upload for asset changes
                 */
                onFileChange: function (oEvent) {
                        const fileUploader = this.byId("fileUploader");
                        const file = fileUploader.oFileUpload.files[0];

                        if (!file) {
                                MessageToast.show("No file selected. Please upload an Excel file with asset changes.");
                                return;
                        }

                        const assetMasterModel = this.getView().getModel("assetMasterEntries");
                        const existingEntries = assetMasterModel ? assetMasterModel.getProperty("/entries") : [];

                        if (existingEntries && existingEntries.length > 0) {
                                MessageBox.confirm(
                                        "Uploading a new file will clear all existing asset change data. Do you want to continue?",
                                        {
                                                title: "Upload New Change File",
                                                onClose: (action) => {
                                                        if (action === MessageBox.Action.OK) {
                                                                this._processFile(file);
                                                        } else {
                                                                fileUploader.clear();
                                                        }
                                                }
                                        }
                                );
                        } else {
                                this._processFile(file);
                        }
                },

                /**
                 * Process uploaded file
                 * @private
                 */
                _processFile: function (file) {
                        console.log("Processing file:", file.name);
                        this._resetData();
                        this.getView().setBusy(true);

                        this._excelProcessor.processExcelFile(file)
                                .then((result) => {
                                        this.getView().setBusy(false);
                                        this._updateUploadSummary(result);

                                        // ✅ NEW: Show specific message based on validation results
                                        if (result.invalidCount > 0) {
                                                MessageToast.show(`File processed with ${result.invalidCount} error(s). Please review before submitting.`);
                                        } else if (result.warningCount > 0) {
                                                MessageToast.show(`File processed successfully with ${result.warningCount} warning(s). Review ValidityStartDate warnings.`);
                                        } else {
                                                MessageToast.show("File processed successfully - all entries valid!");
                                        }
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        console.error("File processing error:", error);
                                        MessageBox.error(`Error processing file: ${error.message || error}`);
                                });
                },

                /**
                 * ✅ ENHANCED: Update upload summary with warnings tracking
                 * @private
                 */
                _updateUploadSummary: function (result) {
                        const uploadSummaryModel = this.getView().getModel("uploadSummary");
                        if (uploadSummaryModel) {
                                uploadSummaryModel.setProperty("/HasBeenSubmitted", false);
                                uploadSummaryModel.setProperty("/IsSubmitEnabled", result.validCount > 0);
                                uploadSummaryModel.setProperty("/ProcessingComplete", true);
                                uploadSummaryModel.setProperty("/WarningCount", result.warningCount || 0);

                                // ✅ NEW: Extract and store all warnings
                                const allWarnings = [];
                                result.entries.forEach(entry => {
                                        if (entry.ValidationWarnings && Array.isArray(entry.ValidationWarnings)) {
                                                entry.ValidationWarnings.forEach(warning => {
                                                        allWarnings.push({
                                                                sequenceId: entry.SequenceNumber,
                                                                asset: entry.FixedAsset,
                                                                companyCode: entry.CompanyCode,
                                                                field: warning.field,
                                                                message: warning.message,
                                                                severity: warning.severity,
                                                                validationType: warning.validationType
                                                        });
                                                });
                                        }
                                });
                                uploadSummaryModel.setProperty("/ValidationWarnings", allWarnings);
                        }
                },

                /**
                 * ✅ ENHANCED: Submit with warning confirmation
                 */
                onSubmitChangesToSAP: function () {
                        const assetMasterModel = this.getView().getModel("assetMasterEntries");
                        const entries = assetMasterModel.getProperty("/entries") || [];
                        const validEntries = entries.filter(entry => entry.Status === 'Valid');

                        if (validEntries.length === 0) {
                                MessageToast.show("No valid asset changes to submit.");
                                return;
                        }

                        // ✅ NEW: Check for entries with warnings
                        const entriesWithWarnings = validEntries.filter(entry =>
                                entry.ValidationWarnings && entry.ValidationWarnings.length > 0
                        );

                        let confirmMessage = `Are you sure you want to submit ${validEntries.length} asset changes to SAP?`;

                        if (entriesWithWarnings.length > 0) {
                                const validityWarnings = entriesWithWarnings.filter(entry =>
                                        entry.ValidationWarnings.some(w => w.validationType === 'VALIDITY_DATE_WARNING')
                                ).length;

                                if (validityWarnings > 0) {
                                        confirmMessage += `\n\n⚠️ WARNING: ${validityWarnings} asset(s) have ValidityStartDate warnings for CHANGE actions.`;
                                        confirmMessage += `\n\nFor CHANGE actions (02), the ValidityStartDate must EXACTLY match an existing interval in SAP.`;
                                        confirmMessage += `\nIf the date doesn't match, SAP will return an error: "time interval doesn't exist".`;
                                        confirmMessage += `\n\nRecommendation: Review the warnings and consider leaving ValidityStartDate blank for CHANGE actions.`;
                                }
                        }

                        MessageBox.confirm(
                                confirmMessage,
                                {
                                        title: "Confirm Asset Changes Submission",
                                        onClose: (action) => {
                                                if (action === MessageBox.Action.OK) {
                                                        this._submitToSAP(validEntries);
                                                }
                                        }
                                }
                        );
                },

                /**
                 * Submit entries to SAP
                 * @private
                 */
                _submitToSAP: function (validEntries) {
                        console.log("Starting SAP submission for", validEntries.length, "asset changes");
                        const entriesToProcess = validEntries.map((entry, index) => ({
                                ...entry,
                                _originalIndex: index,
                                SequenceNumber: entry.SequenceNumber || entry.Sequence || index
                        }));
                        this.getView().setBusy(true);
                        const uploadSummaryModel = this.getView().getModel("uploadSummary");
                        if (uploadSummaryModel) {
                                uploadSummaryModel.setProperty("/isSubmitting", true);
                        }

                        this._batchProcessingManager.processRecordsInBatch(entriesToProcess)
                                .then(result => this._handleSubmissionResult(result))
                                .catch(error => this._handleSubmissionError(error));
                },

                /**
                 * Handle successful submission results
                 * @private
                 */
                _handleSubmissionResult: function (result) {
                        this.getView().setBusy(false);
                        const uploadSummaryModel = this.getView().getModel("uploadSummary");
                        if (uploadSummaryModel) {
                                uploadSummaryModel.setProperty("/isSubmitting", false);
                                uploadSummaryModel.setProperty("/HasBeenSubmitted", true);
                                uploadSummaryModel.setProperty("/ProcessingComplete", true);
                        }

                        const successEntries = result.results ?
                                result.results.filter(r => r.success) : [];
                        const failedEntries = result.results ? result.results.filter(r => !r.success) : [];

                        if (failedEntries.length > 0) {
                                if (successEntries.length > 0) {
                                        this._uiManager.showPartialSuccessWithDocuments(
                                                this._formatEntries(successEntries, 'success'),
                                                this._formatEntries(failedEntries, 'failure')
                                        );
                                } else {
                                        this._uiManager.showErrorWithDetails(
                                                this._formatEntries(failedEntries, 'failure')
                                        );
                                }
                        } else {
                                this._uiManager.showSuccessWithDocuments(
                                        this._formatEntries(successEntries, 'success')
                                );
                        }
                },

                /**
                 * Handle submission errors
                 * @private
                 */
                _handleSubmissionError: function (error) {
                        this.getView().setBusy(false);
                        const uploadSummaryModel = this.getView().getModel("uploadSummary");
                        if (uploadSummaryModel) {
                                uploadSummaryModel.setProperty("/isSubmitting", false);
                        }

                        console.error("Asset change submission error:", error);
                        MessageBox.error(`Failed to submit asset changes: ${error.message || 'Unknown error'}`);
                },

                /**
                 * Format entries for display
                 * @private
                 */
                _formatEntries: function (entries, type) {
                        if (type === 'success') {
                                return entries.map(entry => ({
                                        ...entry.data,
                                        AssetNumber: entry.assetNumber || entry.data?.AssetNumber,
                                        Message: entry.message,
                                        OriginalEntry: entry.originalEntry
                                }));
                        } else {
                                return entries.map(entry => ({
                                        Error: entry.error,
                                        Details: entry.details,
                                        OriginalEntry: entry.originalEntry
                                }));
                        }
                },

                /**
                 * Apply filters with status handling
                 */
                onTableViewChange: function (oEvent) {
                        const selectedKey = oEvent.getSource().getSelectedKey();
                        const filterModel = this.getView().getModel("filterModel");
                        const statusValue = selectedKey === "Valid" ?
                                "valid" :
                                selectedKey === "Invalid" ?
                                        "invalid" : "all";

                        filterModel.setProperty("/status", statusValue);
                        this.applyFilters();

                        const statusComboBox = this.byId("statusFilterComboBox");
                        if (statusComboBox) {
                                statusComboBox.setSelectedKey(statusValue);
                        }
                },

                /**
                 * Enhanced status filter change with sync
                 */
                onStatusFilterChange: function (oEvent) {
                        BaseFilterController.prototype.onStatusFilterChange.apply(this, arguments);
                        const selectedItem = oEvent.getParameter("selectedItem");
                        if (selectedItem) {
                                const selectedKey = selectedItem.getKey();
                                const tableViewSelector = this.byId("tableViewSelector");

                                if (tableViewSelector && selectedKey) {
                                        const segmentKey = selectedKey === "all" ?
                                                "All" :
                                                (selectedKey === "valid" ? "Valid" : "Invalid");
                                        tableViewSelector.setSelectedKey(segmentKey);
                                }
                        }
                },

                /**
                 * Search assets by description
                 */
                onAssetSearch: function (oEvent) {
                        const sQuery = oEvent.getParameter("query");
                        const filterModel = this.getView().getModel("filterModel");

                        filterModel.setData({
                                companyCode: "",
                                assetClass: "",
                                description: sQuery || "",
                                wbsElement: "",
                                status: filterModel.getProperty("/status") || "all",
                                changes: filterModel.getProperty("/changes") || ""
                        });
                        this.applyFilters();
                },

                /**
                 * Download template for asset changes
                 */
                onDownloadTemplate: function () {
                        this.getView().setBusy(true);
                        this._exportManager.downloadChangeTemplate()
                                .then(() => {
                                        this.getView().setBusy(false);
                                        this._errorHandler.showSuccess("Asset change template downloaded successfully");
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        this._errorHandler.showError("Template download failed: " + error.message);
                                });
                },

                /**
                 * ✅ ENHANCED: Show entry details with warnings
                 */
                onShowEntryDetails: function (oEvent) {
                        const oButton = oEvent.getSource();
                        const oBindingContext = oButton.getBindingContext("assetMasterEntries");

                        if (oBindingContext) {
                                const oEntryData = oBindingContext.getObject();

                                // ✅ NEW: Add warnings to entry details if present
                                if (oEntryData.ValidationWarnings && oEntryData.ValidationWarnings.length > 0) {
                                        console.log("Entry has warnings:", oEntryData.ValidationWarnings);
                                }

                                this._uiManager.showEntryDetailsDialog(oEntryData);
                        } else {
                                MessageToast.show("No data found for this asset");
                        }
                },

                /**
                 * Export change processing results
                 */
                onExportResults: function () {
                        console.log("Controller: Starting export results...");
                        const resultsToExport = this._batchProcessingManager?.getResponseData?.();

                        if (!resultsToExport) {
                                MessageToast.show("No change results available to export. Please process some asset changes first.");
                                return;
                        }

                        const hasSuccessRecords = resultsToExport.successRecords && resultsToExport.successRecords.length > 0;
                        const hasErrorRecords = resultsToExport.errorRecords && resultsToExport.errorRecords.length > 0;

                        if (!hasSuccessRecords && !hasErrorRecords) {
                                MessageToast.show("No change results available to export.");
                                return;
                        }

                        this.getView().setBusy(true);
                        this._exportManager.exportChangeResults(resultsToExport, "all", "xlsx")
                                .then(() => {
                                        this.getView().setBusy(false);
                                        MessageToast.show("Change results exported successfully.");
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        console.error("Export failed:", error);
                                        MessageBox.error("Export failed: " + error.message);
                                });
                },

                /**
                 * ✅ ENHANCED: Export validation errors AND warnings
                 */
                onExportChangeValidationErrors: function () {
                        console.log("Controller: Starting export validation errors and warnings...");
                        const assetMasterModel = this.getView().getModel("assetMasterEntries");
                        const entries = assetMasterModel?.getProperty("/entries") || [];

                        const allValidationIssues = [];

                        entries.forEach((entry, index) => {
                                // Collect errors
                                if (entry.ValidationErrors && Array.isArray(entry.ValidationErrors)) {
                                        entry.ValidationErrors.forEach(error => {
                                                allValidationIssues.push({
                                                        sequenceId: entry.SequenceNumber || entry.Sequence || (index + 1),
                                                        field: error.field || error.Field || "Unknown",
                                                        message: error.message || error.Message || "Validation error",
                                                        asset: entry.FixedAsset || entry.AssetNumber || "Unknown",
                                                        severity: "Error",
                                                        type: error.validationType || "VALIDATION_ERROR"
                                                });
                                        });
                                }

                                // ✅ NEW: Collect warnings
                                if (entry.ValidationWarnings && Array.isArray(entry.ValidationWarnings)) {
                                        entry.ValidationWarnings.forEach(warning => {
                                                allValidationIssues.push({
                                                        sequenceId: entry.SequenceNumber || entry.Sequence || (index + 1),
                                                        field: warning.field || warning.Field || "Unknown",
                                                        message: warning.message || warning.Message || "Validation warning",
                                                        asset: entry.FixedAsset || entry.AssetNumber || "Unknown",
                                                        severity: "Warning",
                                                        type: warning.validationType || "VALIDATION_WARNING"
                                                });
                                        });
                                }

                                if (entry.ProcessedValidationErrors && Array.isArray(entry.ProcessedValidationErrors)) {
                                        entry.ProcessedValidationErrors.forEach(error => {
                                                allValidationIssues.push({
                                                        sequenceId: entry.SequenceNumber || entry.Sequence || (index + 1),
                                                        field: error.field || error.Field || "Unknown",
                                                        message: error.message || error.Message || "Validation error",
                                                        asset: entry.FixedAsset || entry.AssetNumber || "Unknown",
                                                        severity: error.severity || "Error",
                                                        type: error.validationType || "VALIDATION_ERROR"
                                                });
                                        });
                                }
                        });

                        // Check uploadSummary model if no issues found
                        if (allValidationIssues.length === 0) {
                                const uploadSummaryModel = this.getView().getModel("uploadSummary");
                                const summaryErrors = uploadSummaryModel?.getProperty("/ValidationErrors") || [];
                                const summaryWarnings = uploadSummaryModel?.getProperty("/ValidationWarnings") || [];

                                [...summaryErrors, ...summaryWarnings].forEach((issue, index) => {
                                        allValidationIssues.push({
                                                sequenceId: issue.sequenceId || issue.SequenceId || (index + 1),
                                                field: issue.field || issue.Field || "Unknown",
                                                message: issue.message || issue.Message || "Validation issue",
                                                asset: issue.asset || issue.Asset || "Unknown",
                                                severity: issue.severity || "Error",
                                                type: issue.validationType || issue.type || "VALIDATION_ISSUE"
                                        });
                                });
                        }

                        if (allValidationIssues.length === 0) {
                                MessageToast.show("No validation errors or warnings to export. Upload a file first.");
                                return;
                        }

                        console.log(`Found ${allValidationIssues.length} validation issues to export (errors and warnings)`);
                        this.getView().setBusy(true);

                        this._exportManager.exportChangeValidationErrors(allValidationIssues)
                                .then(() => {
                                        this.getView().setBusy(false);
                                        MessageToast.show("Validation errors and warnings exported successfully");
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        console.error("Failed to export validation issues:", error);
                                        MessageBox.error("Failed to export validation issues: " + error.message);
                                });
                },

                /**
                 * Export all asset entries (both valid and invalid)
                 */
                onExportAllEntries: function () {
                        console.log("Controller: Starting export all entries...");
                        const assetMasterModel = this.getView().getModel("assetMasterEntries");
                        const entries = assetMasterModel?.getProperty("/entries") || [];

                        if (entries.length === 0) {
                                MessageToast.show("No asset entries to export. Please upload a file first.");
                                return;
                        }

                        this.getView().setBusy(true);
                        const exportData = entries.map((entry, index) => {
                                const processed = this._dataTransformer.processChangeEntryDataForDisplay(entry);
                                return {
                                        SequenceNumber: entry.SequenceNumber || (index + 1),
                                        CompanyCode: entry.CompanyCode || "",
                                        FixedAsset: entry.FixedAsset || "",
                                        Status: entry.Status || "Unknown",
                                        ChangesCount: processed.ChangesSummary?.totalChanges || 0,
                                        HasBlankAssignments: processed.HasBlankAssignments ? "Yes" : "No",
                                        ValidationErrors: entry.ValidationErrors?.length || 0,
                                        ValidationWarnings: entry.ValidationWarnings?.length || 0, // ✅ NEW
                                        ...entry
                                };
                        });

                        const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
                        const fileName = `Asset_Change_Entries_${timestamp}`;

                        this._exportManager.exportToExcel(exportData, fileName, { reportType: "all" })
                                .then(() => {
                                        this.getView().setBusy(false);
                                        MessageToast.show("Asset entries exported successfully");
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        console.error("Failed to export entries:", error);
                                        MessageBox.error("Failed to export entries: " + error.message);
                                });
                },

                /**
                 * Dialog management - unified close method
                 */
                onCloseDialog: function (dialogId) {
                        if (this._uiManager) {
                                this._uiManager.closeDialog(dialogId, true);
                        } else {
                                // Direct fallback
                                const dialog = this.byId(dialogId) ||
                                        sap.ui.getCore().byId(dialogId);
                                if (dialog && typeof dialog.close === 'function') {
                                        dialog.close();
                                }
                        }
                },

                /**
                 * Close batch processing dialog
                 */
                onCloseBatchDialog: function () {
                        if (this._uiManager) {
                                this._uiManager.closeBatchProcessingDialog();
                        } else {
                                this.onCloseDialog("batchProcessingDialog");
                        }
                },

                /**
                 * Cancel batch processing with confirmation
                 */
                onCancelBatchProcessing: function () {
                        MessageBox.confirm(
                                "Are you sure you want to cancel the batch processing?",
                                {
                                        title: "Cancel Batch Processing",
                                        onClose: (action) => {
                                                if (action === MessageBox.Action.OK) {
                                                        // Cancel the processing
                                                        if (this._batchProcessingManager?.cancelProcessing) {
                                                                this._batchProcessingManager.cancelProcessing();
                                                        }

                                                        // Close the dialog
                                                        this.onCloseBatchDialog();
                                                        MessageToast.show("Batch processing cancelled");
                                                }
                                        }
                                }
                        );
                },

                /**
                 * Export handler for batch dialog
                 */
                onExportPress: function (oEvent) {
                        const button = oEvent.getSource();
                        const format = button.data("format") || "XLSX";
                        const type = button.data("type") || "all";

                        console.log(`Export requested - Format: ${format}, Type: ${type}`);

                        // Get results
                        const resultsToExport = this._batchProcessingManager?.getResponseData?.();
                        if (!resultsToExport) {
                                MessageToast.show("No results available to export.");
                                return;
                        }

                        this.getView().setBusy(true);

                        // Export based on format
                        const exportMethod = format.toLowerCase() === 'csv' ?
                                'csv' :
                                format.toLowerCase() === 'pdf' ?
                                        'pdf' : 'xlsx';

                        this._exportManager.exportChangeResults(resultsToExport, type, exportMethod)
                                .then(() => {
                                        this.getView().setBusy(false);
                                        MessageToast.show(`Results exported successfully as ${format}`);
                                })
                                .catch(error => {
                                        this.getView().setBusy(false);
                                        MessageBox.error("Export failed: " + error.message);
                                });
                },

                // Simple specific dialog close handlers
                onChangeValidationDialogClose: function () {
                        this.onCloseDialog("changeValidationErrorsDialog");
                },

                onCloseAssetDetailsDialog: function () {
                        this.onCloseDialog("assetDetailsDialog");
                },

                onValidationDialogClose: function () {
                        this.onCloseDialog("validationErrorsDialog");
                },

                onCloseErrorDialog: function () {
                        this.onCloseDialog("errorDialog");
                },

                /**
                 * Show help dialog
                 */
                onShowHelp: function () {
                        try {
                                this._uiManager.showHelpDialog(this);
                        } catch (error) {
                                console.error("Error showing help dialog:", error);
                                this._errorHandler.showError("Could not display help: " + error.message);
                        }
                },

                /**
                 * Reset application to initial state
                 */
                onReset: function () {
                        MessageBox.confirm(
                                "This will clear all uploaded asset change data and reset the application. Do you want to continue?",
                                {
                                        title: "Reset Application",
                                        onClose: (action) => {
                                                if (action === MessageBox.Action.OK) {
                                                        this._performReset();
                                                }
                                        }
                                }
                        );
                },

                /**
                 * Perform reset operation
                 * @private
                 */
                _performReset: function () {
                        try {
                                this.getView().setBusy(true);
                                this._resetData();
                                this._clearUI();
                                this.getView().setBusy(false);
                                MessageToast.show("Application reset successfully");
                        } catch (error) {
                                this.getView().setBusy(false);
                                console.error("Error during application reset:", error);
                                MessageBox.error("Failed to reset application: " + (error.message || error));
                        }
                },

                /**
                 * Reset data models
                 * @private
                 */
                _resetData: function () {
                        if (this._modelManager) {
                                this._modelManager.resetModels();
                        }
                        this._closeAllDialogs();
                },

                /**
                 * Clear UI elements
                 * @private
                 */
                _clearUI: function () {
                        const fileUploader = this.byId("fileUploader");
                        if (fileUploader) {
                                fileUploader.clear();
                                fileUploader.setValue("");
                        }

                        const filterModel = this.getView().getModel("filterModel");
                        if (filterModel) {
                                filterModel.setData({
                                        companyCode: "",
                                        assetClass: "",
                                        description: "",
                                        wbsElement: "",
                                        status: "all",
                                        changes: ""
                                });
                        }

                        const controls = [
                                { id: "tableViewSelector", method: "setSelectedKey", value: "All" },
                                { id: "statusFilterComboBox", method: "setSelectedKey", value: "all" },
                                { id: "assetSearchField", method: "setValue", value: "" }
                        ];
                        controls.forEach(control => {
                                const element = this.byId(control.id);
                                if (element && typeof element[control.method] === 'function') {
                                        element[control.method](control.value);
                                }
                        });
                },

                /**
                 * Close all dialogs
                 * @private
                 */
                _closeAllDialogs: function () {
                        if (this._uiManager && typeof this._uiManager.closeAllDialogs === 'function') {
                                this._uiManager.closeAllDialogs();
                        } else {
                                // Fallback: close common dialogs directly
                                const dialogIds = [
                                        "batchProcessingDialog",
                                        "changeValidationErrorsDialog",
                                        "assetDetailsDialog",
                                        "validationErrorsDialog",
                                        "errorDialog"
                                ];
                                dialogIds.forEach(id => this.onCloseDialog(id));
                        }
                },

                /**
                 * Utility methods
                 */
                getUIManager: function () {
                        return this._uiManager;
                },

                getBaseURL: function () {
                        const appId = this.getOwnerComponent().getManifestEntry("/sap.app/id");
                        const appPath = appId.replaceAll(".", "/");
                        return jQuery.sap.getModulePath(appPath);
                },

                /**
                 * Formatter method for field changes display
                 */
                formatFieldChangesText: function (fieldChanges) {
                        if (!fieldChanges || !Array.isArray(fieldChanges) || fieldChanges.length === 0) {
                                return "No changes";
                        }

                        if (fieldChanges.length === 1) {
                                return fieldChanges[0].field ||
                                        "1 field";
                        }

                        if (fieldChanges.length <= 3) {
                                return fieldChanges.map(change => change.field || "field").join(", ");
                        }

                        return `${fieldChanges.length} fields changed`;
                },

                /**
                 * Formatter for tooltip
                 */
                formatFieldChangesTooltip: function (fieldChanges) {
                        if (!fieldChanges || !Array.isArray(fieldChanges) || fieldChanges.length === 0) {
                                return "No changes detected";
                        }

                        return fieldChanges.map(change => {
                                const fieldName = change.field || "Unknown field";
                                const newValue = change.isBlankValue ? "<Blank>" : (change.newValue || "");
                                return `${fieldName}: ${newValue}`;
                        }).join("\n");
                },

                /**
                 * Cleanup on exit
                 */
                onExit: function () {
                        console.log("AssetMasterChange Controller: Cleaning up resources...");
                        try {
                                this._closeAllDialogs();
                                if (this._uiManager && typeof this._uiManager.destroy === 'function') {
                                        this._uiManager.destroy();
                                }

                                const modelNames = ["assetMasterEntries", "uploadSummary", "filterModel", "changeSummary", "batchDisplay"];
                                const oView = this.getView();

                                if (oView) {
                                        modelNames.forEach(modelName => {
                                                try {
                                                        const model = oView.getModel(modelName);
                                                        if (model && typeof model.destroy === 'function') {
                                                                model.destroy();
                                                        }
                                                } catch (e) {
                                                        console.warn(`Failed to destroy model ${modelName}:`, e);
                                                }
                                        });
                                }
                        } catch (e) {
                                console.warn("Error during cleanup:", e);
                        }

                        console.log("AssetMasterChange Controller: Cleanup complete");
                }
        });

});