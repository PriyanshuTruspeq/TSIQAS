sap.ui.define([
  "sap/ui/base/Object",
  "sap/ui/model/json/JSONModel"
], function (BaseObject, JSONModel) {
  "use strict";

  return BaseObject.extend("assetmasterchange.model.ModelManager", {

    constructor: function (oController) {
      BaseObject.call(this);
      this._controller = oController;
      this._view = oController.getView();
    },

    /**
     * Initialize all application models for change operations
     */
    initializeModels: function () {
      // ✅ SIMPLIFIED: Initialize core models only
      this._initUploadSummaryModel();
      this._initUserModel();
      this._initAssetMasterModel(); // ✅ CRITICAL: Use correct name
      this._initChangeSummaryModel();
      this._initFilterModel();
      this._initBatchDisplayModel();

      console.log("Asset Change models initialized successfully");
    },

    /**
     * Reset all models to their initial state
     */
    resetModels: function () {
      const models = ["uploadSummary", "assetMasterEntries", "changeSummary", "filterModel", "batchDisplay"];

      models.forEach(modelName => {
        const resetMethod = `_reset${modelName.charAt(0).toUpperCase() + modelName.slice(1)}Model`;
        if (typeof this[resetMethod] === 'function') {
          this[resetMethod]();
        }
      });
    },

    /**
     * ✅ CONSOLIDATED: Single model initialization pattern
     */
    _createModel: function (modelName, data) {
      const model = new JSONModel(data);
      this._view.setModel(model, modelName);
      return model;
    },

    /**
     * ✅ SIMPLIFIED: Upload summary model
     */
    _initUploadSummaryModel: function () {
      this._createModel("uploadSummary", {
        TotalEntries: 0,
        SuccessfulEntries: 0,
        FailedEntries: 0,
        ValidationErrors: [],
        IsSubmitEnabled: false,
        HasBeenSubmitted: false,
        LastUploadDate: null,
        ProcessingComplete: false,
        isSubmitting: false
      });
    },

    /**
     * ✅ SIMPLIFIED: User model
     */
    _initUserModel: function () {
      this._createModel("userInfo", {
        id: "",
        email: "",
        fullName: "",
        displayName: ""
      });
    },

    /**
     * ✅ CRITICAL: Asset master entries model (matches view binding)
     */
    _initAssetMasterModel: function () {
      this.assetMasterModel = this._createModel("assetMasterEntries", {
        entries: [],
        validationStatus: "Pending",
        filteredCount: 0
      });
    },

    /**
     * ✅ SIMPLIFIED: Change summary model
     */
    _initChangeSummaryModel: function () {
      this.changeSummaryModel = this._createModel("changeSummary", {
        totalAssets: 0,
        assetsWithChanges: 0,
        fieldChangeSummary: {},
        blankValueAssignments: []
      });
    },

    /**
     * ✅ SIMPLIFIED: Filter model
     */
    _initFilterModel: function () {
      this._createModel("filterModel", {
        companyCode: "",
        fixedAsset: "",
        description: "",
        wbsElement: "",
        status: "all",
        changes: "all"
      });
    },

    /**
     * ✅ SIMPLIFIED: Batch display model
     */
    _initBatchDisplayModel: function () {
      this._createModel("batchDisplay", {
        visible: false,
        status: "Waiting",
        currentBatch: 0,
        totalBatches: 0,
        processedItems: 0,
        totalItems: 0,
        progressPercentage: 0,
        remainingTime: "Calculating...",
        errors: []
      });
    },

    /**
     * ✅ CONSOLIDATED: Single reset pattern
     */
    _resetModel: function (modelName, data) {
      const model = this._view.getModel(modelName);
      if (model) {
        model.setData(data);
      }
    },

    /**
     * ✅ SIMPLIFIED: Reset methods using consolidated pattern
     */
    _resetUploadSummaryModel: function () {
      this._resetModel("uploadSummary", {
        TotalEntries: 0,
        SuccessfulEntries: 0,
        FailedEntries: 0,
        ValidationErrors: [],
        IsSubmitEnabled: false,
        HasBeenSubmitted: false,
        LastUploadDate: null,
        ProcessingComplete: false,
        isSubmitting: false
      });
    },

    _resetAssetMasterEntriesModel: function () {
      this._resetModel("assetMasterEntries", {
        entries: [],
        validationStatus: "Pending",
        filteredCount: 0
      });
    },

    _resetChangeSummaryModel: function () {
      this._resetModel("changeSummary", {
        totalAssets: 0,
        assetsWithChanges: 0,
        fieldChangeSummary: {},
        blankValueAssignments: []
      });
    },

    _resetFilterModelModel: function () {
      this._resetModel("filterModel", {
        companyCode: "",
        fixedAsset: "",
        description: "",
        wbsElement: "",
        status: "all",
        changes: "all"
      });
    },

    _resetBatchDisplayModel: function () {
      this._resetModel("batchDisplay", {
        visible: false,
        status: "Waiting",
        currentBatch: 0,
        totalBatches: 0,
        processedItems: 0,
        totalItems: 0,
        progressPercentage: 0,
        remainingTime: "Calculating...",
        errors: []
      });
    },

    /**
     * Get model by name
     */
    getModel: function (modelName) {
      return this._view.getModel(modelName);
    },

    /**
     * Update upload summary with validation results
     */
    updateUploadSummary: function (validationResults) {
      const model = this._view.getModel("uploadSummary");
      if (model && validationResults) {
        const updateData = {
          "/TotalEntries": validationResults.entries ? validationResults.entries.length : 0,
          "/SuccessfulEntries": validationResults.validCount || 0,
          "/FailedEntries": validationResults.errorCount || 0,
          "/ValidationErrors": validationResults.errors || [],
          "/IsSubmitEnabled": validationResults.isValid,
          "/LastUploadDate": new Date()
        };

        Object.entries(updateData).forEach(([path, value]) => {
          model.setProperty(path, value);
        });
      }
    },

    /**
     * Update asset change entries
     */
    updateAssetChangeEntries: function (entries, validationStatus = "Pending") {
      const model = this._view.getModel("assetMasterEntries");
      if (model) {
        model.setProperty("/entries", entries || []);
        model.setProperty("/validationStatus", validationStatus);
        model.setProperty("/filteredCount", entries ? entries.length : 0);
      }

      this._updateChangeSummary(entries);
    },

    /**
     * ✅ SIMPLIFIED: Update change summary
     */
    _updateChangeSummary: function (entries) {
      if (!entries || entries.length === 0) {
        this._resetChangeSummaryModel();
        return;
      }

      const summaryData = {
        totalAssets: entries.length,
        assetsWithChanges: 0,
        fieldChangeSummary: {},
        blankValueAssignments: []
      };

      entries.forEach(entry => {
        if (entry.changes && Object.keys(entry.changes).length > 0) {
          summaryData.assetsWithChanges++;

          Object.entries(entry.changes).forEach(([fieldName, changeInfo]) => {
            if (changeInfo.shouldChange) {
              summaryData.fieldChangeSummary[fieldName] = (summaryData.fieldChangeSummary[fieldName] || 0) + 1;

              if (changeInfo.isBlankValue) {
                summaryData.blankValueAssignments.push({
                  asset: entry.FixedAsset,
                  sequenceNumber: entry.SequenceNumber,
                  field: fieldName
                });
              }
            }
          });
        }
      });

      const model = this._view.getModel("changeSummary");
      if (model) {
        model.setData(summaryData);
      }
    },

    /**
     * Update batch processing progress
     */
    updateBatchProgress: function (progressData) {
      const model = this._view.getModel("batchDisplay");
      if (model && progressData) {
        model.setData(progressData);
      }
    },

    /**
     * ✅ SIMPLIFIED: Create validation error model
     */
    createChangeValidationErrorModel: function (errors) {
      const processedErrors = this._processValidationErrors(errors);
      return new JSONModel(processedErrors);
    },

    /**
     * ✅ CONSOLIDATED: Process validation errors
     */
    _processValidationErrors: function (errors) {
      const flatErrors = [];
      const affectedAssets = new Set();
      const errorCategories = {};

      errors.forEach(error => {
        const processedError = {
          message: error.message || error.Message || 'Unknown error',
          field: error.field || error.Field || 'Unknown',
          severity: error.severity || error.Severity || 'Error',
          sequenceId: error.sequenceId || error.SequenceId || 'N/A',
          asset: error.asset || error.AssetNumber || error.FixedAsset || 'Unknown',
          category: this._categorizeError(error)
        };

        flatErrors.push(processedError);
        affectedAssets.add(processedError.asset);

        if (!errorCategories[processedError.category]) {
          errorCategories[processedError.category] = [];
        }
        errorCategories[processedError.category].push(processedError);
      });

      return {
        flatErrors,
        errorCount: flatErrors.length,
        affectedAssets: affectedAssets.size,
        errorCategories
      };
    },

    /**
     * ✅ SIMPLIFIED: Categorize errors
     */
    _categorizeError: function (error) {
      const message = (error.message || '').toLowerCase();

      if (message.includes('required') || message.includes('missing')) {
        return 'Required Field';
      } else if (message.includes('blank') || message.includes('clear')) {
        return 'Blank Assignment';
      } else if (message.includes('format') || message.includes('invalid')) {
        return 'Format Error';
      } else if (message.includes('length') || message.includes('exceed')) {
        return 'Length Validation';
      } else {
        return 'General Validation';
      }
    }
  });
});