sap.ui.define([
  "sap/ui/model/json/JSONModel",
  "sap/ui/core/Fragment"
], function (JSONModel, Fragment) {
  "use strict";

  /**
   * UIManager for AssetMasterChange - CLEANED VERSION
   * Centralized dialog management including batch processing
   */
  return function (oController, oErrorHandlerInstance, oExportManagerInstance) {

    // Validate dependencies
    if (!oController) {
      throw new Error("UIManager requires a controller instance.");
    }
    if (!oErrorHandlerInstance) {
      throw new Error("UIManager requires a valid ErrorHandler instance.");
    }
    if (!oExportManagerInstance) {
      throw new Error("UIManager requires a valid ExportManager instance.");
    }

    this.oController = oController;
    this.errorHandler = oErrorHandlerInstance;
    this.exportManager = oExportManagerInstance;
    this.dataTransformer = oController._dataTransformer;
    this._dialogs = {};
    this._helpControllerContext = null;

    /**
     * Unified dialog loading and display
     * @private
     */
    this._loadAndShowDialog = function (dialogId, fragmentName, oModel, modelName, oControllerOverride) {
      const openDialog = (oDialog) => {
        if (!oDialog || oDialog.bIsDestroyed) {
          console.error(`UIManager: Invalid dialog: ${dialogId}`);
          this.errorHandler.showError("Could not open dialog window.");
          delete this._dialogs[dialogId];
          return;
        }

        if (oModel) {
          oDialog.setModel(oModel, modelName || undefined);
        }

        if (typeof oDialog.isOpen === 'function' && !oDialog.isOpen()) {
          oDialog.open();
        } else if (typeof oDialog.open === 'function') {
          oDialog.open();
        }
      };

      // Handle problematic dialogs that need recreation
      if (["assetDetailsDialog", "validationErrorsDialog", "batchProcessingDialog"].includes(dialogId)) {
        this._destroyExistingDialog(dialogId);
        return new Promise((resolve) => {
          setTimeout(() => {
            this._loadDialogFragment(dialogId, fragmentName, oModel, modelName, oControllerOverride, openDialog)
              .then(resolve)
              .catch((error) => {
                console.error(`UIManager: Failed to load dialog: ${dialogId}`, error);
                throw error;
              });
          }, 100);
        });
      }

      // Check cache for normal dialogs
      let oExistingDialog = this._dialogs[dialogId];
      if (oExistingDialog && !oExistingDialog.bIsDestroyed) {
        if (sap.ui.getCore().byId(oExistingDialog.getId())) {
          openDialog(oExistingDialog);
          return Promise.resolve(oExistingDialog);
        } else {
          delete this._dialogs[dialogId];
        }
      }

      return this._loadDialogFragment(dialogId, fragmentName, oModel, modelName, oControllerOverride, openDialog);
    };

    /**
     * Destroy existing dialog safely
     * @private
     */
    this._destroyExistingDialog = function (dialogId) {
      let existingDialog = this._dialogs[dialogId];

      if (!existingDialog || existingDialog.bIsDestroyed) {
        const sFragmentInstanceId = this.oController.getView().createId(dialogId);
        existingDialog = sap.ui.getCore().byId(sFragmentInstanceId) ||
          sap.ui.getCore().byId(dialogId);
      }

      if (existingDialog && !existingDialog.bIsDestroyed) {
        try {
          const parent = existingDialog.getParent();
          if (parent && typeof parent.removeDependent === 'function') {
            parent.removeDependent(existingDialog);
          }
          existingDialog.destroy();
          console.log(`UIManager: Destroyed existing ${dialogId}`);
        } catch (e) {
          console.error(`UIManager: Error destroying ${dialogId}:`, e);
        }
      }

      delete this._dialogs[dialogId];
    };

    /**
     * Load dialog fragment
     * @private
     */
    this._loadDialogFragment = function (dialogId, fragmentName, oModel, modelName, oControllerOverride, openDialog) {
      const timestamp = Date.now();
      const uniqueFragmentId = `${this.oController.getView().getId()}--${dialogId}-${timestamp}`;

      return new Promise((resolve, reject) => {
        Fragment.load({
          id: uniqueFragmentId,
          name: fragmentName,
          controller: oControllerOverride || this.oController
        }).then((oDialog) => {
          if (!oDialog) {
            reject(new Error("Fragment.load resolved without a valid control instance."));
            return;
          }

          const dialogInstance = Array.isArray(oDialog) ? oDialog[0] : oDialog;
          if (!dialogInstance || typeof dialogInstance.addDependent !== 'function') {
            reject(new Error("Loaded fragment root is not a valid control instance."));
            return;
          }

          this._dialogs[dialogId] = dialogInstance;
          this.oController.getView().addDependent(dialogInstance);
          console.log(`UIManager: Fragment loaded: ${dialogId}`);

          openDialog(dialogInstance);
          resolve(dialogInstance);

        }).catch((error) => {
          console.error(`UIManager: Could not load fragment: ${fragmentName}`, error);
          this.errorHandler.showError(`Could not load dialog: ${fragmentName}`, error.message);
          delete this._dialogs[dialogId];
          reject(error);
        });
      });
    };

    /**
     * Show any dialog by ID and fragment name
     * @public
     */
    this.showDialog = function (dialogId, fragmentName, oModel, modelName, oControllerOverride) {
      console.log(`UIManager: showDialog called for ${dialogId}`);
      return this._loadAndShowDialog(dialogId, fragmentName, oModel, modelName, oControllerOverride);
    };

    /**
        * Simplified close dialog method
        * @public
        */
    this.closeDialog = function (dialogId, bDestroy = false) {
      if (!dialogId) {
        console.warn("UIManager: No dialog ID provided");
        return;
      }

      console.log(`UIManager: Closing dialog: ${dialogId}`);

      let oDialog = this._dialogs[dialogId];

      // If not in cache, try to find it
      if (!oDialog || oDialog.bIsDestroyed) {
        // Try with view ID prefix first
        const viewId = this.oController.getView().getId();
        oDialog = sap.ui.getCore().byId(`${viewId}--${dialogId}`) ||
          sap.ui.getCore().byId(dialogId) ||
          this.oController.byId(dialogId);
      }

      if (oDialog && typeof oDialog.close === 'function' && !oDialog.bIsDestroyed) {
        try {
          // Simple close - let UI5 handle the state checking
          oDialog.close();
          console.log(`UIManager: Dialog closed: ${dialogId}`);

          // Handle destruction if requested or for specific dialogs
          if (bDestroy || this._shouldAutoDestroy(dialogId)) {
            this._destroyDialog(oDialog, dialogId);
          }
        } catch (error) {
          console.error(`UIManager: Error closing dialog ${dialogId}:`, error);
        }
      } else {
        console.warn(`UIManager: Dialog not found or already destroyed: ${dialogId}`);
      }
    };

    /**
     * Check if dialog should be auto-destroyed
     * @private
     */
    this._shouldAutoDestroy = function (dialogId) {
      // Auto-destroy these dialogs to prevent stale data
      const autoDestroyDialogs = [
        "assetDetailsDialog",
        "batchProcessingDialog",
        "validationErrorsDialog"
      ];
      return autoDestroyDialogs.includes(dialogId);
    };

    /**
     * Simplified dialog destruction
     * @private
     */
    this._destroyDialog = function (oDialog, dialogId) {
      try {
        if (oDialog && !oDialog.bIsDestroyed && typeof oDialog.destroy === 'function') {
          // Remove from parent if needed
          const parent = oDialog.getParent();
          if (parent && typeof parent.removeDependent === 'function') {
            parent.removeDependent(oDialog);
          }

          oDialog.destroy();
          console.log(`UIManager: Dialog destroyed: ${dialogId}`);
        }

        // Clean up cache
        delete this._dialogs[dialogId];
      } catch (error) {
        console.error(`UIManager: Error destroying dialog ${dialogId}:`, error);
        // Still clean up cache even if destruction failed
        delete this._dialogs[dialogId];
      }
    };

    /**
     * Simplified close batch processing dialog
     * @public
     */
    this.closeBatchProcessingDialog = function () {
      this.closeDialog("batchProcessingDialog", true);
    };

    /**
     * Simplified close all dialogs
     * @public
     */
    this.closeAllDialogs = function () {
      console.log("UIManager: Closing all dialogs...");

      // Close all cached dialogs
      Object.keys(this._dialogs).forEach(dialogId => {
        this.closeDialog(dialogId, true);
      });

      // Clear the cache
      this._dialogs = {};
    };

    /**
     * Simplified destroy method
     * @public
     */
    this.destroy = function () {
      console.log("UIManager: Destroying UIManager...");

      // Close and destroy all dialogs
      this.closeAllDialogs();

      // Clean up references
      this._helpControllerContext = null;
      this.oController = null;
      this.errorHandler = null;
      this.exportManager = null;
      this.dataTransformer = null;
    };

    /**
     * Show batch processing progress dialog
     * @public
     */
    this.showBatchProcessingDialog = function (totalEntries, batchOptions = {}) {
      console.log("UIManager: Showing batch processing dialog...");

      const batchDisplayData = {
        status: "Initializing batch processing...",
        totalEntries: totalEntries,
        processedEntries: 0,
        successCount: 0,
        failureCount: 0,
        percentage: 0,
        timeRemaining: "Calculating...",
        processingSpeed: "0 entries/sec",
        currentBatch: 0,
        totalBatches: Math.ceil(totalEntries / (batchOptions.batchSize || 10)),
        isCompleted: false,
        isError: false,
        batchOptions: batchOptions
      };

      const oBatchModel = new JSONModel(batchDisplayData);

      // Set the model on the view FIRST before loading dialog
      this.oController.getView().setModel(oBatchModel, "batchDisplay");
      console.log("UIManager: batchDisplay model set on view with data:", batchDisplayData);

      return this._loadAndShowDialog(
        "batchProcessingDialog",
        "assetmasterchange.view.BatchProcessingDisplayDialog",
        oBatchModel,
        "batchDisplay",
        this._createBatchDialogController()
      ).then((dialog) => {
        // Ensure model is also set on dialog
        if (dialog && typeof dialog.setModel === 'function') {
          dialog.setModel(oBatchModel, "batchDisplay");
          console.log("UIManager: batchDisplay model also set on dialog");
        }
        return dialog;
      });
    };

    /**
 * Update batch processing progress
 * @public
 */
    this.updateBatchProgress = function (progressData) {
      // Check view model first
      let batchModel = this.oController.getView().getModel("batchDisplay");

      if (!batchModel) {
        console.warn("UIManager: batchDisplay model not found on view, trying to create one");
        // Create a new model if it doesn't exist
        batchModel = new JSONModel({
          status: "Processing...",
          totalEntries: 0,
          processedEntries: 0,
          successCount: 0,
          failureCount: 0,
          percentage: 0,
          timeRemaining: "Calculating...",
          processingSpeed: "0 entries/sec",
          currentBatch: 0,
          totalBatches: 0,
          isCompleted: false,
          isError: false
        });
        this.oController.getView().setModel(batchModel, "batchDisplay");
        console.log("UIManager: Created new batchDisplay model");
      }

      if (batchModel && progressData) {
        console.log("UIManager: Updating batch progress with:", progressData);

        // Get current data to merge with new data
        const currentData = batchModel.getData();
        const mergedData = { ...currentData, ...progressData };

        // Update the entire model data
        batchModel.setData(mergedData);

        // Also update individual properties for better reactivity
        Object.keys(progressData).forEach(key => {
          const newValue = progressData[key];
          batchModel.setProperty(`/${key}`, newValue);
          console.log(`UIManager: Updated ${key}: ${newValue}`);
        });

        // Force model refresh
        batchModel.refresh(true);
        console.log("UIManager: Model refreshed with merged data:", mergedData);
      } else {
        console.warn("UIManager: No batch model or progress data available", {
          hasModel: !!batchModel,
          hasProgressData: !!progressData
        });
      }
    };

    /**
     * Finalize batch processing (completion/error state)
     * @public
     */
    this.finalizeBatchProgress = function (success, error, finalData = {}) {
      const batchModel = this.oController.getView().getModel("batchDisplay");
      if (!batchModel) return;

      const currentData = batchModel.getData();
      if (currentData && currentData.isCompleted) return;

      const finalStatus = success ? "Processing completed successfully" :
        (error ? `Processing completed with errors: ${error.message || error}` : "Processing completed with errors");

      batchModel.setData({
        ...currentData,
        ...finalData,
        status: finalStatus,
        error: error && !success ? (error.message || error) : "",
        timeRemaining: "Completed",
        isCompleted: true,
        isError: !success
      });
    };

    /**
     * Create batch dialog controller with handlers
     * @private
     */
 
    this._createBatchDialogController = function () {
      console.log("UIManager: Creating batch dialog controller");

      return {
        // Handler for Close button
        onCloseBatchDialog: () => {
          console.log("UIManager BatchController: onCloseBatchDialog called");
          this.closeDialog("batchProcessingDialog", true);
        },

        // Handler for Cancel button  
        onCancelBatchProcessing: () => {
          console.log("UIManager BatchController: onCancelBatchProcessing called");

          // Show confirmation dialog
          sap.m.MessageBox.confirm(
            "Are you sure you want to cancel the batch processing?",
            {
              title: "Cancel Batch Processing",
              onClose: (action) => {
                if (action === sap.m.MessageBox.Action.OK) {
                  // Cancel the processing
                  if (this.oController._batchProcessingManager &&
                    typeof this.oController._batchProcessingManager.cancelProcessing === 'function') {
                    this.oController._batchProcessingManager.cancelProcessing();
                  }

                  // Close the dialog
                  this.closeDialog("batchProcessingDialog", true);
                  sap.m.MessageToast.show("Batch processing cancelled");
                }
              }
            }
          );
        },

        // Handler for Export buttons
        onExportPress: (oEvent) => {
          console.log("UIManager BatchController: onExportPress called");

          try {
            const button = oEvent.getSource();
            const format = button.data("format") || "XLSX";
            const type = button.data("type") || "all";

            console.log(`UIManager BatchController: Export requested - Format: ${format}, Type: ${type}`);

            // Get results from batch processing manager
            const resultsToExport = this.oController._batchProcessingManager?.getResponseData?.();
            if (!resultsToExport) {
              sap.m.MessageToast.show("No results available to export.");
              return;
            }

            // Show busy indicator on the main view
            this.oController.getView().setBusy(true);

            // Export based on format
            const exportMethod = format.toLowerCase() === 'csv' ? 'csv' :
              format.toLowerCase() === 'pdf' ? 'pdf' : 'xlsx';

            this.exportManager.exportChangeResults(resultsToExport, type, exportMethod)
              .then(() => {
                this.oController.getView().setBusy(false);
                sap.m.MessageToast.show(`Results exported successfully as ${format}`);
              })
              .catch(error => {
                this.oController.getView().setBusy(false);
                console.error("UIManager BatchController: Export failed:", error);
                sap.m.MessageBox.error("Export failed: " + error.message);
              });

          } catch (error) {
            this.oController.getView().setBusy(false);
            console.error("UIManager BatchController: Error in onExportPress:", error);
            sap.m.MessageBox.error("Export failed: " + error.message);
          }
        },

        // Alternative handler name (in case fragment uses different name)
        onBatchProcessingDialogClose: () => {
          console.log("UIManager BatchController: onBatchProcessingDialogClose called");
          this.closeDialog("batchProcessingDialog", true);
        }
      };
    };

    /**
     * Close batch processing dialog specifically
     * @public
     */
    this.closeBatchProcessingDialog = function () {
      this.closeDialog("batchProcessingDialog", true);
    };

    /**
     * Handle validation errors
     * @public
     */
    this.handleValidationErrors = function (errors) {
      console.log("UIManager: Handling validation errors...");

      if (!errors || !errors.length) return;

      if (!this.dataTransformer || typeof this.dataTransformer.processValidationErrorsForDisplay !== 'function') {
        console.error("DataTransformer not available for error processing");
        return;
      }

      const processedErrors = this.dataTransformer.processValidationErrorsForDisplay(errors);
      if (!processedErrors.flatErrors || !processedErrors.flatErrors.length) return;

      const oValidationModel = new JSONModel(processedErrors);
      this._loadAndShowDialog(
        "validationErrorsDialog",
        "assetmasterchange.view.ValidationErrorsDialog",
        oValidationModel,
        "validation",
        this.oController
      );
    };

    /**
     * Show change entry details
     * @public
     */
    this.showChangeEntryDetailsDialog = function (oEntry) {
      console.log("UIManager: Showing change entry details...");

      if (!oEntry) {
        this.errorHandler.showWarning("Cannot show details: No change entry data provided.");
        return;
      }

      if (!this.dataTransformer || typeof this.dataTransformer.processChangeEntryDataForDisplay !== 'function') {
        this.errorHandler.showError("DataTransformer not available for processing change entry data.");
        return;
      }

      const processedEntryData = this.dataTransformer.processChangeEntryDataForDisplay(oEntry);
      this._loadAndShowDialog(
        "assetDetailsDialog",
        "assetmasterchange.view.EntryDetailsDialog",
        new JSONModel(processedEntryData),
        "changeDetails",
        this.oController
      );
    };

    // Alias for backward compatibility
    this.showEntryDetailsDialog = this.showChangeEntryDetailsDialog;

    /**
     * Unified result dialog display
     * @private
     */
    this._showResultDialog = function (dialogType, entries, secondaryEntries) {
      const dialogConfigs = {
        success: {
          id: "successDialog",
          fragment: "assetmasterchange.view.ChangeSuccessDialog",
          formatter: "formatChangeSuccessEntriesForDisplay",
          model: "success"
        },
        partial: {
          id: "partialSuccessDialog",
          fragment: "assetmasterchange.view.PartialChangeSuccessDialog",
          formatter: "formatPartialChangeResultsForDisplay",
          model: "partial"
        },
        error: {
          id: "errorDialog",
          fragment: "assetmasterchange.view.ChangeErrorDialog",
          formatter: "formatFailedChangeEntriesForDisplay",
          model: "error"
        }
      };

      const config = dialogConfigs[dialogType];
      if (!config) return;

      if (dialogType === 'success' && (!entries || !entries.length)) {
        this.errorHandler.showSuccess("Asset changes completed successfully!");
        return;
      }

      if ((dialogType === 'error' || dialogType === 'partial') && (!entries || !entries.length)) {
        if (dialogType === 'partial') {
          this.errorHandler.showWarning("No asset changes were processed.");
        }
        return;
      }

      let formattedData;
      if (dialogType === 'partial') {
        formattedData = this.exportManager[config.formatter](entries, secondaryEntries);
      } else {
        formattedData = this.exportManager[config.formatter](entries);
      }

      const oModel = new JSONModel(formattedData);
      const controllerOverride = this._createDialogController(config.id);

      this._loadAndShowDialog(config.id, config.fragment, oModel, config.model, controllerOverride);
    };

    /**
     * Create dialog controller with close handler
     * @private
     */
    this._createDialogController = function (dialogId) {
      const closeMethodName = `on${dialogId.charAt(0).toUpperCase() + dialogId.slice(1)}Close`;

      return {
        [closeMethodName]: () => {
          if (this.oController && typeof this.oController[closeMethodName] === 'function') {
            this.oController[closeMethodName]();
          } else {
            this.closeDialog(dialogId, true);
          }
        }
      };
    };

    // Public result dialog methods
    this.showChangeSuccessWithDetails = function (successEntries) {
      this._showResultDialog('success', successEntries);
    };

    this.showPartialChangeSuccessWithDetails = function (successEntries, failedEntries) {
      this._showResultDialog('partial', successEntries, failedEntries);
    };

    this.showChangeErrorWithDetails = function (failedEntries) {
      this._showResultDialog('error', failedEntries);
    };

    // Aliases for backward compatibility
    this.showSuccessWithDocuments = this.showChangeSuccessWithDetails;
    this.showPartialSuccessWithDocuments = this.showPartialChangeSuccessWithDetails;
    this.showErrorWithDetails = this.showChangeErrorWithDetails;

    /**
     * Show help dialog
     * @public
     */
    this.showHelpDialog = function (controllerContext) {
      console.log("UIManager: Showing help dialog...");
      this._helpControllerContext = controllerContext && typeof controllerContext.onDownloadTemplate === 'function' ? controllerContext : null;

      this._loadAndShowDialog("helpDialog", "assetmasterchange.view.HelpDialog", null, null, {
        onHelpDialogClose: () => {
          this.closeDialog("helpDialog", false);
          this._helpControllerContext = null;
        },
        onDownloadChangeTemplateLinkPress: () => {
          if (this._helpControllerContext) {
            this._helpControllerContext.onDownloadTemplate();
          } else {
            this.errorHandler.showWarning("Cannot download change template: Application context not available.");
          }
        }
      });
    };

    /**
     * Export validation errors
     * @public
     */
    this.exportValidationErrors = function (groupedErrorsData) {
      return this.exportManager.exportValidationErrorsToCSV(groupedErrorsData);
    };

    /**
     * Close all open dialogs
     * @public
     */
    this.closeAllDialogs = function () {
      console.log("UIManager: Closing all dialogs...");

      // Close cached dialogs
      Object.keys(this._dialogs).forEach(dialogId => {
        try {
          this.closeDialog(dialogId, true);
        } catch (e) {
          console.error(`Error closing cached dialog ${dialogId}:`, e);
        }
      });

      // Search for any remaining dialogs manually
      this._searchAndCloseDialogs();
    };

    /**
     * Search UI5 registry for open dialogs
     * @private
     */
    this._searchAndCloseDialogs = function () {
      try {
        const allControls = sap.ui.getCore().mElements || {};

        Object.keys(allControls).forEach(controlId => {
          try {
            const control = allControls[controlId];

            if (control &&
              !control.bIsDestroyed &&
              control.getMetadata &&
              typeof control.getMetadata === 'function') {

              const metadataName = control.getMetadata().getName();

              if ((metadataName.includes("Dialog") || metadataName.includes("Popover")) &&
                typeof control.close === 'function') {

                if (typeof control.isOpen === 'function') {
                  if (control.isOpen()) {
                    control.close();
                    console.log(`UIManager: Found and closed dialog: ${controlId}`);
                  }
                } else {
                  control.close();
                  console.log(`UIManager: Found and closed control: ${controlId}`);
                }
              }
            }
          } catch (e) {
            // Ignore individual control errors
          }
        });
      } catch (error) {
        console.error("UIManager: Error searching for dialogs:", error);
      }
    };

    /**
     * Clean up all dialogs and resources
     * @public
     */
    this.destroy = function () {
      console.log("UIManager: Destroying all dialogs...");

      Object.keys(this._dialogs).forEach(dialogId => {
        try {
          const dialog = this._dialogs[dialogId];
          if (dialog && !dialog.bIsDestroyed && typeof dialog.destroy === 'function') {
            dialog.destroy();
            console.log(`UIManager: Destroyed dialog ${dialogId}`);
          }
        } catch (e) {
          console.error(`UIManager: Error destroying dialog ${dialogId}:`, e);
        }
      });

      this._dialogs = {};
      this._helpControllerContext = null;
    };
  };
});