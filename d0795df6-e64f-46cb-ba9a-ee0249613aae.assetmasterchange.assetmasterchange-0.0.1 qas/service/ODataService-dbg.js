sap.ui.define(
    [
        "sap/ui/model/odata/v4/ODataModel",
        "sap/ui/model/json/JSONModel",
        "sap/m/MessageBox",
        "sap/ui/core/BusyIndicator",
        "sap/ui/model/Filter",
        "sap/ui/model/FilterOperator",
        "sap/m/MessageToast",
        "sap/ui/core/format/DateFormat",
        "sap/ui/core/Core",
        "sap/ui/model/odata/v4/Context",
        "sap/ui/core/message/Message",
        "assetmasterchange/utils/ErrorUtils"
    ],
    function (
        ODataModel,
        JSONModel,
        MessageBox,
        BusyIndicator,
        Filter,
        FilterOperator,
        MessageToast,
        DateFormat,
        Core,
        Context,
        Message,
        ErrorUtils
    ) {
        "use strict";

        /**
         * ODataService class for handling communication with OData services
         * Enhanced with automatic ValidityStartDate reading for CHANGE operations
         */
        return class ODataService {
            /**
             * Constructor for ODataService
             * @param {sap.ui.model.odata.v4.ODataModel} oModel - OData v4 model
             * @param {object} dataTransformer - Instance of the DataTransformer utility
             * @param {object} errorHandler - Instance of the ErrorHandler utility
             */
            constructor(oModel, dataTransformer, errorHandler) {
                if (!oModel) {
                    throw new Error("OData model is required");
                }
                if (!dataTransformer) {
                    throw new Error("DataTransformer is required");
                }
                if (!errorHandler) {
                    throw new Error("ErrorHandler is required");
                }

                const modelName = oModel.getMetadata ? oModel.getMetadata().getName() : null;
                if (modelName !== "sap.ui.model.odata.v4.ODataModel") {
                    throw new Error(`Invalid OData model type. Expected sap.ui.model.odata.v4.ODataModel, got ${modelName}`);
                }

                this._oModel = oModel;
                this._dataTransformer = dataTransformer;
                this._errorHandler = errorHandler;
                this._basePath = "/FixedAsset";
                this._dateFormatter = DateFormat.getDateInstance({ pattern: "yyyy-MM-dd" });
                this._batchUpdateGroupId = "assetBatchGroup";
                this._debugMode = true;

                this._resetInternalState();
                this._initializeODataErrorHandlers();
            }

            /**
             * Initialize error handlers for the OData model
             * @private
             */
            _initializeODataErrorHandlers() {
                try {
                    const oModel = this._oModel;
                    if (oModel && oModel.getMetadata().getName() === "sap.ui.model.odata.v4.ODataModel") {
                        const originalReportError = oModel.reportError;

                        oModel.reportError = (sLogMessage, sReportingClassName, oError) => {
                            if (this._preventEntityRefresh &&
                                oError &&
                                oError.message &&
                                oError.message.includes("No key predicate known")) {
                                console.warn("Ignoring expected OData refresh error during batch processing:", oError.message);
                                return;
                            }
                            return originalReportError.call(oModel, sLogMessage, sReportingClassName, oError);
                        };

                        console.log("OData error handlers initialized");
                    }
                } catch (error) {
                    console.warn("Could not initialize OData error handlers:", error);
                }
            }

            /**
             * Reset internal state of the service
             * @private
             */
            _resetInternalState() {
                this.totalRecords = 0;
                this.successCount = 0;
                this.failureCount = 0;
                this.failedRecords = [];
                this.successEntries = [];
                this._isCancelled = false;
                this._isProcessing = false;
                this._totalBatches = 0;
                this._batchStartTime = null;
                this._preventEntityRefresh = false;
                this._processedCount = 0;
                this._currentSequenceGroups = {};
            }

            /**
             * Set debug mode for detailed logging
             * @param {boolean} enable - Whether to enable debug mode
             */
            setDebugMode(enable) {
                this._debugMode = Boolean(enable);
                this._logInfo(`Debug mode ${this._debugMode ? 'enabled' : 'disabled'}`);
            }

            /**
             * Enhanced entity validation that also retrieves ETag for If-Match header
             * @private
             */
            _validateEntityExists(companyCode, assetNumber, subNumber) {
                const entityPath = `/FixedAsset/${companyCode}/${assetNumber}/${subNumber}`;
                this._logInfo(`Validating entity exists and retrieving ETag at: ${entityPath}`);

                return new Promise((resolve, reject) => {
                    try {
                        const entityBinding = this._oModel.bindContext(entityPath);

                        entityBinding.requestObject("$direct").then(
                            (data) => {
                                this._logInfo(`Entity validation successful for ${assetNumber}:`, data);

                                let etag = null;
                                try {
                                    const context = entityBinding.getBoundContext();
                                    if (context && context.getObject) {
                                        const contextData = context.getObject();
                                        etag = contextData["@odata.etag"] ||
                                            context.getProperty("@odata.etag") ||
                                            entityBinding.getModel().getETag(entityPath);
                                    }
                                } catch (etagError) {
                                    this._logWarn(`Could not retrieve ETag for ${assetNumber}:`, etagError);
                                }

                                if (!etag) {
                                    etag = "*";
                                    this._logInfo(`Using wildcard ETag for ${assetNumber}`);
                                } else {
                                    this._logInfo(`Retrieved ETag for ${assetNumber}: ${etag}`);
                                }

                                resolve({ data, etag });
                            },
                            (error) => {
                                this._logError(`Entity validation failed for ${assetNumber}:`, error);
                                reject(new Error(`Asset not found: ${assetNumber} in company ${companyCode}. Details: ${error.message}`));
                            }
                        );
                    } catch (error) {
                        this._logError(`Entity validation setup failed:`, error);
                        reject(error);
                    }
                });
            }

            /**
             * Read existing asset data including ledger, valuation, and time-based valuation
             * @private
             */
            _readExistingAssetData(companyCode, assetNumber, subNumber) {
                const entityPath = `/FixedAsset(CompanyCode='${companyCode}',MasterFixedAsset='${assetNumber}',FixedAsset='${subNumber}')`;
                const expandQuery = "$expand=_Ledger($expand=_Valuation($expand=_TimeBasedValuation))";
                const fullPath = `${entityPath}?${expandQuery}`;

                this._logInfo(`Reading existing asset data from: ${fullPath}`);

                return new Promise((resolve, reject) => {
                    try {
                        const binding = this._oModel.bindContext(fullPath);

                        binding.requestObject("$direct").then(
                            (data) => {
                                this._logInfo(`Successfully read existing asset data for ${assetNumber}`);
                                resolve(data);
                            },
                            (error) => {
                                this._logWarn(`Could not read existing asset data for ${assetNumber}:`, error);
                                // Don't reject - missing data means CREATE operation
                                resolve(null);
                            }
                        );
                    } catch (error) {
                        this._logError(`Error setting up read for existing asset data:`, error);
                        resolve(null); // Don't reject - allow operation to continue
                    }
                });
            }

            /**
             * Extract existing ValidityStartDate from SAP data
             * @private
             */
            _extractExistingValidityStartDate(existingData, ledgerCode, areaCode) {
                if (!existingData || !existingData._Ledger) {
                    return null;
                }

                try {
                    // Find matching ledger
                    const ledger = existingData._Ledger.find(l => l.Ledger === ledgerCode);
                    if (!ledger || !ledger._Valuation) {
                        return null;
                    }

                    // Find matching valuation/depreciation area
                    const valuation = ledger._Valuation.find(v => v.AssetDepreciationArea === areaCode);
                    if (!valuation || !valuation._TimeBasedValuation || valuation._TimeBasedValuation.length === 0) {
                        return null;
                    }

                    // Get the first (earliest) time-based valuation entry
                    const timeBasedValuation = valuation._TimeBasedValuation[0];
                    const validityStartDate = timeBasedValuation.ValidityStartDate;

                    if (validityStartDate) {
                        this._logInfo(`Found existing ValidityStartDate for ${ledgerCode}-${areaCode}: ${validityStartDate}`);
                        return validityStartDate;
                    }

                    return null;
                } catch (error) {
                    this._logError(`Error extracting ValidityStartDate:`, error);
                    return null;
                }
            }

            /**
             * Suspend OData operations to prevent issues during batch processing
             */
            suspendODataOperations() {
                try {
                    const oModel = this._oModel;
                    if (!oModel) return;

                    if (oModel.getMetadata().getName() === "sap.ui.model.odata.v4.ODataModel") {
                        this._modelState = {
                            autoRefresh: oModel.isAutoRefreshEnabled ? oModel.isAutoRefreshEnabled() : false
                        };

                        if (oModel.setAutoRefreshEnabled) {
                            oModel.setAutoRefreshEnabled(false);
                        }
                        if (oModel.suspend) {
                            oModel.suspend();
                        }

                        this._preventEntityRefresh = true;
                        console.log("OData operations suspended");
                    }
                } catch (error) {
                    console.warn("Error suspending OData operations:", error);
                }
            }

            /**
             * Resume OData operations after batch processing
             */
            resumeODataOperations() {
                try {
                    const oModel = this._oModel;
                    if (!oModel) return;

                    if (oModel.getMetadata().getName() === "sap.ui.model.odata.v4.ODataModel") {
                        if (this._modelState && oModel.setAutoRefreshEnabled) {
                            oModel.setAutoRefreshEnabled(this._modelState.autoRefresh);
                        }
                        if (oModel.resume) {
                            oModel.resume();
                        }

                        this._preventEntityRefresh = false;
                        console.log("OData operations resumed");
                    }
                } catch (error) {
                    console.warn("Error resuming OData operations:", error);
                }
            }

            /**
             * Change asset masters grouped by sequence ID
             * @param {Array} assetChanges - Array of asset change entries to process
             * @param {Object} callbacks - Callback functions for batch processing
             * @returns {Promise} Promise resolving with processing results
             */
            changeAssetMasters(assetChanges, callbacks = {}) {
                if (!assetChanges || !Array.isArray(assetChanges) || assetChanges.length === 0) {
                    const error = new Error("Valid asset change entries array is required");
                    if (callbacks.error) {
                        callbacks.error({ message: error.message });
                    }
                    return Promise.reject(error);
                }

                if (this._isProcessing) {
                    const error = new Error("Another batch process is already running");
                    if (callbacks.error) {
                        callbacks.error({ message: error.message });
                    }
                    return Promise.reject(error);
                }

                try {
                    this._resetInternalState();
                    this._isProcessing = true;
                    this.totalRecords = assetChanges.length;
                    this._batchStartTime = new Date();

                    this._currentSequenceGroups = this._groupBySequenceId(assetChanges);
                    this._totalBatches = Object.keys(this._currentSequenceGroups).length;

                    this.suspendODataOperations();

                    this._logInfo(`Starting processing of ${this.totalRecords} asset changes grouped into ${this._totalBatches} sequence batches`);

                    return this._processSequenceGroups(this._currentSequenceGroups, callbacks);
                } catch (initError) {
                    this.resumeODataOperations();
                    this._isProcessing = false;
                    this._logError("Error initializing batch processing", initError);

                    if (callbacks.error) {
                        callbacks.error({
                            message: "Error initializing batch processing: " + initError.message
                        });
                    }
                    return Promise.reject(initError);
                }
            }

            /**
             * Group asset changes by sequence ID
             * @private
             */
            _groupBySequenceId(assetChanges) {
                const groups = {};

                assetChanges.forEach((asset, index) => {
                    const sequenceId = asset.Sequence || asset.OriginalSequence || asset.SequenceNumber || `NoSequence_${index}`;

                    if (!groups[sequenceId]) {
                        groups[sequenceId] = [];
                    }
                    asset._originalIndex = index;
                    groups[sequenceId].push(asset);
                });

                this._logInfo(`Grouped ${assetChanges.length} assets into ${Object.keys(groups).length} sequence groups`);

                Object.keys(groups).forEach(sequenceId => {
                    this._logInfo(`Sequence ${sequenceId}: ${groups[sequenceId].length} items`);
                });

                return groups;
            }

            /**
             * Process all sequence groups
             * @private
             */
            _processSequenceGroups(sequenceGroups, callbacks) {
                return new Promise((resolve) => {
                    const sequenceIds = Object.keys(sequenceGroups);

                    const processNextSequenceGroup = (sequenceIndex) => {
                        if (this._isCancelled) {
                            this._logInfo("Processing cancelled by user");
                            this._isProcessing = false;
                            this.resumeODataOperations();
                            const result = this._createFinalResult(true);
                            if (callbacks.success) {
                                callbacks.success(result);
                            }
                            resolve(result);
                            return;
                        }

                        if (sequenceIndex >= sequenceIds.length) {
                            this._logInfo("All sequence groups processed successfully");
                            this._isProcessing = false;
                            this.resumeODataOperations();
                            const result = this._createFinalResult();
                            if (callbacks.success) {
                                try {
                                    callbacks.success(result);
                                } catch (callbackError) {
                                    this._logError("Error in success callback", callbackError);
                                }
                            }
                            resolve(result);
                            return;
                        }

                        const currentSequenceId = sequenceIds[sequenceIndex];
                        const sequenceEntries = sequenceGroups[currentSequenceId];

                        this._logInfo(`Processing sequence group ${sequenceIndex + 1}/${sequenceIds.length}: ${currentSequenceId} with ${sequenceEntries.length} entries`);

                        if (callbacks.batchStart) {
                            try {
                                callbacks.batchStart(sequenceIndex + 1, sequenceIds.length);
                            } catch (error) {
                                this._logError("Error in batchStart callback", error);
                            }
                        }

                        this._processSequenceBatch(sequenceEntries, currentSequenceId, sequenceIndex, sequenceIds.length, callbacks)
                            .then(() => {
                                if (this._isCancelled) {
                                    processNextSequenceGroup(sequenceIds.length);
                                    return;
                                }
                                setTimeout(() => {
                                    processNextSequenceGroup(sequenceIndex + 1);
                                }, 300);
                            })
                            .catch(error => {
                                this._logError(`Error processing sequence group ${currentSequenceId}`, error);
                                if (this._isCancelled) {
                                    processNextSequenceGroup(sequenceIds.length);
                                    return;
                                }
                                setTimeout(() => {
                                    processNextSequenceGroup(sequenceIndex + 1);
                                }, 500);
                            });
                    };

                    processNextSequenceGroup(0);
                });
            }

            /**
             * Process a sequence batch using individual requests
             * @private
             */
            _processSequenceBatch(sequenceEntries, sequenceId, sequenceIndex, totalSequences, callbacks) {
                this._logInfo(`Processing sequence batch: ${sequenceId} with ${sequenceEntries.length} entries`);

                return new Promise((resolve) => {
                    const processNextEntry = async (entryIndex) => {
                        if (this._isCancelled) {
                            resolve();
                            return;
                        }

                        if (entryIndex >= sequenceEntries.length) {
                            this._updateProgress(sequenceIndex, totalSequences, callbacks);
                            this._logInfo(`Completed sequence batch: ${sequenceId}`);
                            resolve();
                            return;
                        }

                        const entry = sequenceEntries[entryIndex];
                        const entryId = entry._originalIndex || entryIndex;

                        try {
                            const result = await this._processSingleAssetChange(entry);

                            if (result.success) {
                                this.successCount++;
                                this._processedCount++;
                                this.successEntries.push({
                                    ...result.response,
                                    OriginalRequest: entry,
                                    OriginalSequence: sequenceId,
                                    Status: "Success",
                                    ProcessedAt: new Date().toISOString(),
                                    Message: result.message || `Asset change completed successfully in sequence ${sequenceId}`,
                                    success: true
                                });
                            } else {
                                this.failureCount++;
                                this._processedCount++;
                                this.failedRecords.push({
                                    entry: entry,
                                    error: result.error || "Unknown error occurred",
                                    details: result.details || result.error,
                                    originalSequence: sequenceId,
                                    errorCode: result.errorCode
                                });
                            }

                            this._updateProgress(sequenceIndex, totalSequences, callbacks);

                        } catch (error) {
                            this._logError(`Error processing entry ${entryId} in sequence ${sequenceId}`, error);
                            this.failureCount++;
                            this._processedCount++;
                            this.failedRecords.push({
                                entry: entry,
                                error: ErrorUtils.extractErrorMessage(error) || "Unknown error occurred",
                                details: error.message,
                                originalSequence: sequenceId
                            });

                            this._updateProgress(sequenceIndex, totalSequences, callbacks);
                        }

                        setTimeout(() => {
                            processNextEntry(entryIndex + 1);
                        }, 100);
                    };

                    processNextEntry(0);
                });
            }

            /**
             * Process a single asset change with automatic ValidityStartDate reading
             * @private
             */
            _processSingleAssetChange(assetEntry) {
                const entryId = assetEntry._originalIndex ?? assetEntry.SequenceNumber ?? "unknown";
                this._logInfo(`Processing asset change entry index ${entryId}`);

                return new Promise(async (resolve) => {
                    try {
                        const companyCode = assetEntry.CompanyCode || assetEntry.CompCode || "";
                        const assetNumber = assetEntry.FixedAsset || assetEntry.AssetNumber || "";
                        const subNumber = assetEntry.FixedAssetSubNumber || assetEntry.SubNumber || "0";

                        if (!companyCode || !assetNumber) {
                            throw new Error(`Missing required parameters: CompanyCode='${companyCode}', AssetNumber='${assetNumber}'`);
                        }

                        // Step 1: Validate entity exists and get ETag
                        let entityData, etag;
                        try {
                            const validationResult = await this._validateEntityExists(companyCode, assetNumber, subNumber);
                            entityData = validationResult.data;
                            etag = validationResult.etag;
                            this._logInfo(`Entity validation passed for ${assetNumber} with ETag: ${etag}`);
                        } catch (validationError) {
                            this._logError(`Entity validation failed for ${assetNumber}:`, validationError);
                            resolve({
                                success: false,
                                error: `Asset not found: ${assetNumber}`,
                                errorCode: "ENTITY_NOT_FOUND",
                                details: { validationError: validationError.message },
                                originalInput: assetEntry
                            });
                            return;
                        }

                        // Step 2: Read existing asset data for CHANGE operations
                        let existingAssetData = null;
                        try {
                            existingAssetData = await this._readExistingAssetData(companyCode, assetNumber, subNumber);
                            if (existingAssetData) {
                                this._logInfo(`Retrieved existing asset data for ${assetNumber}`);
                            }
                        } catch (readError) {
                            this._logWarn(`Could not read existing asset data for ${assetNumber}:`, readError);
                            // Continue - this just means we can't auto-fill dates for CHANGE operations
                        }

                        // Step 3: Transform the payload with existing data context
                        const transformedPayload = this._dataTransformer.transformToODataFormat(assetEntry, existingAssetData);
                        if (!transformedPayload || Object.keys(transformedPayload).length === 0) {
                            throw new Error(`Data transformation returned empty result for entry ${entryId}. Check DataTransformer.`);
                        }

                        this._logInfo(`Transformed payload for asset change index ${entryId}:`, JSON.stringify(transformedPayload, null, 2));

                        // Step 4: Clear previous messages
                        Core.getMessageManager().removeAllMessages();

                        // Step 5: Execute the change operation
                        try {
                            const result = await this._executeChangeUsingDirectCallWithETag(
                                companyCode, assetNumber, subNumber,
                                transformedPayload, etag, entryId, assetEntry
                            );
                            this._logInfo(`Change operation succeeded for ${entryId}:`, result);
                            resolve(result);
                        } catch (executeError) {
                            this._logError(`Change operation failed for ${entryId}:`, executeError);
                            resolve({
                                success: false,
                                error: executeError.message || "Change operation failed",
                                errorCode: "EXECUTION_FAILED",
                                details: { lastError: executeError.message, payload: transformedPayload, etag },
                                originalInput: assetEntry
                            });
                        }

                    } catch (setupError) {
                        this._logError(`Setup error for index ${entryId}`, setupError);
                        resolve({
                            success: false,
                            error: setupError.message || "Setup failed",
                            errorCode: "SETUP_ERROR",
                            details: { originalError: setupError },
                            originalInput: assetEntry
                        });
                    }
                });
            }

            /**
             * Direct HTTP call with If-Match header
             * @private
             */
            _executeChangeUsingDirectCallWithETag(companyCode, assetNumber, subNumber, transformedPayload, etag, entryId, assetEntry) {
                console.log("🚀 DIRECT HTTP CALL WITH ETAG METHOD!", entryId);
                this._logInfo(`USING DIRECT HTTP CALL WITH ETAG METHOD for ${entryId}, ETag: ${etag}`);

                return new Promise((resolve, reject) => {
                    try {
                        const serviceUrl = this._oModel.sServiceUrl;
                        const changeUrl = `${serviceUrl}FixedAsset(CompanyCode='${companyCode}',MasterFixedAsset='${assetNumber}',FixedAsset='${subNumber}')/SAP__self.Change`;

                        this._logInfo(`Direct call URL with ETag: ${changeUrl}`);
                        this._logInfo(`Payload for direct call:`, JSON.stringify(transformedPayload, null, 2));
                        this._logInfo(`Using If-Match header: ${etag}`);

                        const xhr = new XMLHttpRequest();
                        xhr.open('POST', changeUrl, true);

                        xhr.setRequestHeader('Content-Type', 'application/json');
                        xhr.setRequestHeader('Accept', 'application/json');
                        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                        xhr.setRequestHeader('If-Match', etag);

                        const headers = this._oModel.getHttpHeaders ? this._oModel.getHttpHeaders() : {};
                        const csrfToken = headers['X-CSRF-Token'] || headers['x-csrf-token'];
                        if (csrfToken) {
                            xhr.setRequestHeader('X-CSRF-Token', csrfToken);
                        }

                        xhr.onload = () => {
                            if (xhr.status >= 200 && xhr.status < 300) {
                                this._logInfo(`Direct call with ETag successful for ${entryId}. Status: ${xhr.status}`);

                                let response;
                                try {
                                    response = JSON.parse(xhr.responseText);
                                } catch (parseError) {
                                    response = {
                                        CompanyCode: companyCode,
                                        MasterFixedAsset: assetNumber,
                                        FixedAsset: subNumber,
                                        message: "Change completed successfully",
                                        success: true
                                    };
                                }

                                resolve({
                                    success: true,
                                    response: response,
                                    message: response.message || `Asset change completed successfully for ${assetNumber}`,
                                    code: "SUCCESS",
                                    originalInput: assetEntry,
                                    sentPayload: transformedPayload,
                                    usedETag: etag
                                });
                            } else {
                                this._logError(`Direct call with ETag failed for ${entryId}. Status: ${xhr.status}, Response: ${xhr.responseText}`);

                                let errorMessage = `HTTP ${xhr.status}`;
                                try {
                                    const errorResponse = JSON.parse(xhr.responseText);
                                    errorMessage = errorResponse.error?.message || errorResponse.message || errorMessage;
                                } catch (parseError) {
                                    errorMessage = xhr.responseText || errorMessage;
                                }

                                reject(new Error(`OData Call Error: ${errorMessage}`));
                            }
                        };

                        xhr.onerror = () => {
                            this._logError(`Direct call with ETag network error for ${entryId}`);
                            reject(new Error('Network error during direct call with ETag'));
                        };

                        xhr.ontimeout = () => {
                            this._logError(`Direct call with ETag timeout for ${entryId}`);
                            reject(new Error('Timeout during direct call with ETag'));
                        };

                        xhr.timeout = 30000;

                        xhr.send(JSON.stringify(transformedPayload));

                    } catch (setupError) {
                        this._logError(`Direct call with ETag setup error for ${entryId}:`, setupError);
                        reject(setupError);
                    }
                });
            }

            /**
             * Update progress of batch processing
             * @private
             */
            _updateProgress(sequenceIndex, totalSequences, callbacks) {
                if (callbacks && callbacks.batchProgress && typeof callbacks.batchProgress === 'function') {
                    try {
                        const processedCount = this.successCount + this.failureCount;
                        this._logInfo(`Progress update: Sequence ${sequenceIndex + 1}/${totalSequences}, Processed: ${processedCount}/${this.totalRecords}`);

                        callbacks.batchProgress(
                            sequenceIndex + 1,
                            totalSequences,
                            processedCount,
                            this.totalRecords
                        );
                    } catch (error) {
                        this._logError("Error in batchProgress callback", error);
                    }
                }
            }

            /**
             * Create final result object
             * @private
             */
            _createFinalResult(cancelled = false) {
                return {
                    cancelled: cancelled,
                    totalRecords: this.totalRecords,
                    successCount: this.successCount,
                    failureCount: this.failureCount,
                    successRecords: this.successEntries,
                    errorRecords: this.failedRecords,
                    successfulRecords: this.successEntries,
                    failedRecordsList: this.failedRecords
                };
            }

            /**
             * Cancel the current batch processing
             */
            cancelBatchProcessing() {
                if (this._isProcessing) {
                    this._isCancelled = true;
                    this._logInfo("Cancellation requested for batch processing");
                    this._isProcessing = false;

                    return Promise.resolve({
                        cancelled: true,
                        totalRecords: this.totalRecords,
                        successCount: this.successCount,
                        failureCount: this.failureCount,
                        failedRecords: this.failedRecords,
                        successRecords: this.successEntries,
                        failedRecordsList: this.failedRecords,
                        successfulRecords: this.successEntries,
                        responseData: {
                            successItems: this.successCount,
                            errorItems: this.failedRecords
                        }
                    });
                }

                return Promise.resolve({
                    cancelled: false,
                    message: "No active processing to cancel"
                });
            }

            /**
             * Check if batch processing is currently running
             */
            isProcessing() {
                return this._isProcessing;
            }

            /**
             * Log info message to console
             * @private
             */
            _logInfo(message, details) {
                if (this._debugMode) {
                    console.log(`[ODataService] INFO: ${message}`, details !== undefined ? details : "");
                }
            }

            /**
             * Log warning message to console
             * @private
             */
            _logWarn(message, details) {
                console.warn(`[ODataService] WARN: ${message}`, details !== undefined ? details : "");
            }

            /**
             * Log error message to console
             * @private
             */
            _logError(message, error) {
                console.error(`[ODataService] ERROR: ${message}`, error !== undefined ? error : "");
            }
        };
    }
);