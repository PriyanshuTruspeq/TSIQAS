sap.ui.define([
    "sap/ui/model/json/JSONModel",
    "assetmasterchange/utils/ProgressTracker",
    "assetmasterchange/utils/ErrorUtils"
], function (JSONModel, ProgressTracker, ErrorUtils) {
    "use strict";

    return function (options) {
        // Store options and controller reference
        this.oController = options.controller;
        this._uiManager = options.uiManager;
        this._errorHandler = options.errorHandler;
        this._dataTransformer = options.dataTransformer;
        this._oDataService = options.oDataService;
        this._exportManager = options.ExportManager;

        this._progressTracker = new ProgressTracker();
        this._timeoutId = null;
        this._batchOptions = {};

        // Processing state
        this._processingState = {
            isProcessing: false,
            isCanceled: false,
            currentBatchIndex: 0,
            totalBatches: 0
        };

        // Response data
        this._responseData = {
            totalRecords: 0,
            successCount: 0,
            failureCount: 0,
            successRecords: [],
            errorRecords: [],
            allMessages: [],
            initialEntries: []
        };

        this._logInfo = function (message) {
            console.log(`[BatchProcessingManager] ${message}`);
        };

        this._logError = function (message, error) {
            console.error(`[BatchProcessingManager] ${message}`, error || "");
        };

        /**
         * Initialize batch processing display via UIManager
         * @private
         */
        this._initBatchProcessingDisplay = function (totalEntries) {
            if (!this._batchOptions.showProgress) {
                return Promise.resolve();
            }

            return this._uiManager.showBatchProcessingDialog(totalEntries, this._batchOptions)
                .then(() => {
                    this._logInfo("Batch processing dialog initialized via UIManager");
                })
                .catch(error => {
                    this._logError("Error initializing batch display via UIManager", error);
                    // Continue processing even if dialog fails
                });
        };

        /**
         * Update progress display via UIManager
         * @private
         */
        this._updateProgressDisplay = function () {
            if (!this._batchOptions.showProgress) return;

            try {
                const progress = this._progressTracker.getProgress();
                const elapsedMinutes = progress.elapsedTime / 60;
                const speed = elapsedMinutes > 0 ? Math.round(progress.processed / elapsedMinutes) : 0;

                const progressData = {
                    totalEntries: progress.total,
                    processedEntries: progress.processed,
                    successCount: progress.successCount,
                    failureCount: progress.failureCount,
                    percentage: progress.percentage,
                    timeRemaining: progress.timeRemaining,
                    currentBatch: this._processingState.currentBatchIndex,
                    totalBatches: this._processingState.totalBatches,
                    processingSpeed: `${speed} entries/min`,
                    entriesProgress: progress.percentage,
                    sequenceProgress: this._processingState.totalBatches > 0 ?
                        (this._processingState.currentBatchIndex / this._processingState.totalBatches) * 100 : 0,
                    displayText: {
                        entries: `${progress.processed} of ${progress.total} entries`,
                        sequences: `${this._processingState.currentBatchIndex} of ${this._processingState.totalBatches} sequence groups`
                    }
                };

                console.log("BatchProcessingManager: Calling updateBatchProgress with:", progressData);
                this._uiManager.updateBatchProgress(progressData);
            } catch (error) {
                this._logError("Error updating progress display", error);
            }
        };

        /**
         * Finalize progress via UIManager
         * @private
         */
        this._finalizeProgress = function (success, error) {
            if (!this._batchOptions.showProgress) return;

            try {
                const progress = this._progressTracker.getProgress();

                const finalData = {
                    timeRemaining: "Completed",
                    processingTime: `${progress.elapsedTime.toFixed(2)}s`,
                    processedEntries: progress.processed,
                    successCount: progress.successCount,
                    failureCount: progress.failureCount,
                    percentage: progress.percentage
                };

                this._uiManager.finalizeBatchProgress(success, error, finalData);
            } catch (e) {
                this._logError("Error finalizing progress", e);
            }
        };

        /**
         * Main processing method
         * @public
         */
        this.processRecordsInBatch = function (entries, options = {}) {
            if (!entries || !entries.length) {
                return Promise.reject(new Error("No entries to process"));
            }

            if (!this._oDataService) {
                return Promise.reject(new Error("OData Service not initialized"));
            }

            // Store batch options
            this._batchOptions = {
                batchSize: 10,
                showProgress: true,
                ...options
            };

            // Initialize state
            this._processingState = {
                isProcessing: true,
                isCanceled: false,
                currentBatchIndex: 0,
                totalBatches: Math.ceil(entries.length / this._batchOptions.batchSize)
            };
            if (this._processingState.totalBatches === 0 && entries.length > 0) {
                this._processingState.totalBatches = 1;
            }

            // Initialize progress tracker
            this._progressTracker.start(entries.length);

            // Initialize response data
            this._responseData = {
                totalRecords: entries.length,
                successCount: 0,
                failureCount: 0,
                successRecords: [],
                errorRecords: [],
                allMessages: [],
                initialEntries: entries.map(entry => ({ ...entry }))
            };

            if (this._timeoutId) {
                clearTimeout(this._timeoutId);
                this._timeoutId = null;
            }

            const initPromise = this._batchOptions.showProgress
                ? this._initBatchProcessingDisplay(entries.length)
                : Promise.resolve();

            return initPromise.then(() => {
                const callbacks = {
                    batchStart: (sequenceGroupIndex, totalSequenceGroups) => {
                        this._logInfo(`ODataService: Sequence group ${sequenceGroupIndex} of ${totalSequenceGroups} initiated.`);

                        this._uiManager.updateBatchProgress({
                            status: `Processing sequence group ${sequenceGroupIndex} of ${totalSequenceGroups}...`
                        });

                        this._processingState.currentBatchIndex = sequenceGroupIndex;
                        this._processingState.totalBatches = totalSequenceGroups;
                        this._updateProgressDisplay();
                    },

                    batchProgress: (sequenceGroupIndex, totalSequenceGroups, processedItems, totalItems) => {
                        this._uiManager.updateBatchProgress({
                            status: `Processing sequence group ${sequenceGroupIndex} of ${totalSequenceGroups} - ${processedItems} of ${totalItems} items processed...`,
                            processedEntries: processedItems,
                            percentage: totalItems > 0 ? (processedItems / totalItems) * 100 : 0
                        });

                        this._processingState.currentBatchIndex = sequenceGroupIndex;
                        if (this._processingState.totalBatches !== totalSequenceGroups) {
                            this._processingState.totalBatches = totalSequenceGroups;
                        }
                        this._updateProgressDisplay();
                    },

                    success: (result) => {
                        this._logInfo("ODataService: Sequence-based batch processing reported success.");
                        this._uiManager.updateBatchProgress({
                            status: "Processing completed successfully"
                        });
                        this._handleBatchResult(result);
                        this._finalizeProgress(true);
                        return this._createSuccessResponse(result);
                    },

                    error: (error, resultFromODataErrorCallback) => {
                        this._logError("ODataService: Sequence-based batch processing reported an error.", error);
                        this._uiManager.updateBatchProgress({
                            status: "Processing completed with errors"
                        });

                        if (resultFromODataErrorCallback) {
                            this._handleBatchResult(resultFromODataErrorCallback);
                        } else {
                            const totalFailureResult = {
                                totalRecords: this._responseData.totalRecords,
                                successCount: 0,
                                failureCount: this._responseData.totalRecords,
                                successRecords: [],
                                failedRecordsList: this._responseData.initialEntries.map(entry => ({
                                    originalEntry: entry,
                                    Sequence: entry.Sequence || entry.SequenceNumber || "N/A",
                                    Status: "Error",
                                    RawData: entry,
                                    ErrorMessage: `Asset master creation failed: ${error.message || "Unknown error"}`,
                                    ErrorDetails: error.stack || error.details,
                                    ValidationErrors: [],
                                    error: { message: `Asset master creation failed: ${error.message || "Unknown error"}` }
                                }))
                            };
                            this._handleBatchResult(totalFailureResult);
                        }
                        this._finalizeProgress(false, error);
                        return this._createErrorResponse(error);
                    }
                };

                return this._oDataService.changeAssetMasters(entries, callbacks);
            });
        };

        /**
         * Handle batch result processing
         * @private
         */
        this._handleBatchResult = function (result) {
            if (!result) return;

            if (result.successRecords) {
                const successRecords = result.successRecords || [];
                this._responseData.successCount += successRecords.length;
                this._responseData.successRecords = [...this._responseData.successRecords, ...successRecords];
                this._progressTracker.update(successRecords.length, true, successRecords);
            }

            const errorRecordsList = result.errorRecords || result.failedRecordsList || [];
            if (errorRecordsList.length > 0) {
                this._responseData.failureCount += errorRecordsList.length;
                this._responseData.errorRecords = [...this._responseData.errorRecords, ...errorRecordsList];
                this._progressTracker.update(errorRecordsList.length, false, errorRecordsList);
            }

            if (result.error || result.cancelled) {
                this._processingState.isProcessing = false;
                if (result.cancelled) {
                    this._processingState.isCanceled = true;
                }
            }
            this._updateProgressDisplay();
        };

        /**
         * Create success response
         * @private
         */
        this._createSuccessResponse = function (result) {
            const progress = this._progressTracker.getProgress();
            return {
                success: true,
                successCount: progress.successCount,
                errorCount: progress.failureCount,
                totalCount: progress.total,
                results: [
                    ...this._responseData.successRecords.map(record => ({
                        success: true,
                        data: record,
                        message: record.Message || "Asset created successfully",
                        assetNumber: record.AssetNumber || record.FixedAsset,
                        originalEntry: record.OriginalRequest
                    })),
                    ...this._responseData.errorRecords.map(record => ({
                        success: false,
                        error: record.error || record.ErrorMessage,
                        details: record.details || record.ErrorDetails,
                        originalEntry: record.entry || record.originalEntry
                    }))
                ]
            };
        };

        /**
         * Create error response
         * @private
         */
        this._createErrorResponse = function (error) {
            const progress = this._progressTracker.getProgress();
            return {
                error: true,
                errorMessage: error ? error.message : "Unknown error occurred",
                successCount: progress.successCount,
                errorCount: progress.failureCount,
                totalCount: progress.total,
                results: [
                    ...this._responseData.successRecords.map(record => ({
                        success: true,
                        data: record,
                        message: record.Message || "Asset created successfully",
                        assetNumber: record.AssetNumber || record.FixedAsset,
                        originalEntry: record.OriginalRequest
                    })),
                    ...this._responseData.errorRecords.map(record => ({
                        success: false,
                        error: record.error || record.ErrorMessage,
                        details: record.details || record.ErrorDetails,
                        originalEntry: record.entry || record.originalEntry
                    }))
                ]
            };
        };

        /**
         * Cancel processing
         * @public
         */
        this.cancelProcessing = function () {
            if (this._processingState.isProcessing && !this._processingState.isCanceled) {
                this._processingState.isCanceled = true;

                if (this._oDataService?.cancelBatchProcessing) {
                    try {
                        this._oDataService.cancelBatchProcessing();
                        this._logInfo("Cancellation signal sent to ODataService.");
                    } catch (error) {
                        this._logError("Error signaling cancel to ODataService", error);
                    }
                }

                this._uiManager.updateBatchProgress({
                    status: "Cancelling processing..."
                });
                return true;
            }
            return false;
        };

        /**
         * Close batch processing dialog via UIManager
         * @public
         */
        this.closeBatchProcessingDialog = function () {
            try {
                this._uiManager.closeBatchProcessingDialog();
            } catch (error) {
                this._logError("Error closing batch processing dialog via UIManager", error);
            }
        };

        /**
         * Get response data
         * @public
         */
        this.getResponseData = function () {
            const progress = this._progressTracker.getProgress();
            return {
                ...this._responseData,
                successCount: progress.successCount,
                failureCount: progress.failureCount,
                totalRecords: progress.total
            };
        };

        /**
         * Get processing state
         * @public
         */
        this.getProcessingState = function () {
            return this._processingState;
        };

        /**
         * Force close dialog and reset state
         * @public
         */
        this.forceCloseDialog = function () {
            this.closeBatchProcessingDialog();
            this._processingState.isProcessing = false;
            if (this._timeoutId) {
                clearTimeout(this._timeoutId);
                this._timeoutId = null;
            }
        };

        // Legacy method name for compatibility
        this.startBatchProcessing = function (entries, options = {}) {
            return this.processRecordsInBatch(entries, options);
        };
    };
});