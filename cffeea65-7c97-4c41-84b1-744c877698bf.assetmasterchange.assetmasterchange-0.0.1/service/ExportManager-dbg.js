sap.ui.define([
  "assetmasterchange/utils/ErrorHandler",
  "assetmasterchange/utils/DataTransformer",
  "assetmasterchange/utils/XlsxFormatter",
  "assetmasterchange/utils/CsvFormatter",
  "assetmasterchange/utils/SapSpreadsheetFormatter",
  "assetmasterchange/utils/PdfFormatter",
  "assetmasterchange/utils/TemplateGenerator"
], function (
  ErrorHandler,
  DataTransformer,
  XlsxFormatter,
  CsvFormatter,
  SapSpreadsheetFormatter,
  PdfFormatter,
  TemplateGenerator
) {
  "use strict";

  // Helper function to trigger browser download from a Blob
  const triggerDownload = (blob, fileName) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100); // Cleanup
  };

  return class ExportManager {
    /**
     * Constructor - Inject dependencies
     * @param {object} [options={}] - Constructor options
     * @param {sap.ui.core.mvc.Controller} [options.controller] - Controller instance (optional)
     * @param {ErrorHandler} [options.errorHandler] - Error handler instance (optional)
     * @param {object} [options.customExcelProcessor] - Optional custom processor for templates
     */
    constructor(options = {}) {
      // Instantiate dependencies
      this.oController = options.controller;
      this._errorHandler = options.errorHandler || new ErrorHandler();
      this._dataTransformer = options.dataTransformer || new DataTransformer();
      this._xlsxFormatter = new XlsxFormatter();
      this._csvFormatter = new CsvFormatter();
      this._sapSpreadsheetFormatter = new SapSpreadsheetFormatter();
      this._templateGenerator = new TemplateGenerator({
        xlsxFormatter: this._xlsxFormatter,
        customExcelProcessor: options.customExcelProcessor
      });
      this._pdfFormatter = new PdfFormatter();
    }

    // =================
    // DATA PROCESSING METHODS (delegated to DataTransformer)
    // =================

    /**
     * Processes validation errors for display in dialogs.
     * @param {Array} errors - Array of validation errors
     * @returns {Object} Object with flatErrors, errorCount, and _rawGroupedData
     * @public
     */
    processValidationErrorsForDisplay(errors) {
      return this._dataTransformer.processValidationErrorsForDisplay(errors);
    }

    /**
     * Processes entry data for display in details dialog.
     * @param {Object} oEntry - Entry data object
     * @returns {Object} Processed entry data
     * @public
     */
    processEntryDataForDisplay(oEntry) {
      return this._dataTransformer.processChangeEntryDataForDisplay(oEntry);
    }

    /**
     * Calculates validation statistics.
     * @param {number} total - Total count
     * @param {number} successful - Successful count
     * @param {number} failed - Failed count
     * @returns {Object} Statistics object
     * @public
     */
    calculateValidationStats(total = 0, successful = 0, failed = 0) {
      return this._dataTransformer.calculateValidationStats(total, successful, failed);
    }

    // =================
    // CHANGE-SPECIFIC FORMATTING METHODS
    // =================

    /**
     * Format change success entries for display
     * @param {Array} successEntries - Array of successful change entries
     * @returns {Object} Formatted data for display
     */
    formatChangeSuccessEntriesForDisplay(successEntries) {
      console.log("ExportManager: Formatting change success entries for display...");

      if (!successEntries || !Array.isArray(successEntries)) {
        return {
          entries: [],
          summary: {
            totalCount: 0,
            successfulChanges: 0
          }
        };
      }

      const processedEntries = successEntries.map((entry, index) => {
        const processedEntry = this._dataTransformer.processChangeEntryDataForDisplay(entry);
        return {
          ...processedEntry,
          index: index + 1,
          status: "Success",
          assetNumber: entry.AssetNumber || entry.FixedAsset || entry.MasterFixedAsset || `Asset-${index + 1}`
        };
      });

      return {
        entries: processedEntries,
        summary: {
          totalCount: successEntries.length,
          successfulChanges: successEntries.length,
          timestamp: new Date().toISOString()
        }
      };
    }

    /**
     * Format change error entries for display  
     * @param {Array} failedEntries - Array of failed change entries
     * @returns {Object} Formatted data for display
     */
    formatFailedChangeEntriesForDisplay(failedEntries) {
      console.log("ExportManager: Formatting failed change entries for display...");

      if (!failedEntries || !Array.isArray(failedEntries)) {
        return {
          entries: [],
          summary: {
            totalCount: 0,
            failedChanges: 0
          }
        };
      }

      const processedEntries = failedEntries.map((entry, index) => {
        // Handle the new error response structure
        let errorMessage = "";
        let errorCode = "UNKNOWN_ERROR";
        let errorDetails = "";

        if (typeof entry.error === 'string') {
          errorMessage = entry.error;
        } else if (entry.error) {
          errorMessage = entry.error.message || "Unknown error";
          errorCode = entry.error.code || errorCode;
          errorDetails = entry.details || "";

          // Handle the SAP specific error format
          if (entry.error.innererror) {
            const innerError = entry.error.innererror;
            if (innerError.ErrorDetails) {
              errorDetails += JSON.stringify(innerError.ErrorDetails);
            }
          }
        } else {
          errorMessage = entry.Error || entry.error || "Unknown error";
          errorDetails = entry.Details || entry.details || "";
          errorCode = entry.ErrorCode || entry.errorCode || errorCode;
        }

        return {
          index: index + 1,
          status: "Failed",
          error: errorMessage,
          details: errorDetails,
          assetNumber: entry.assetNumber || entry.AssetNumber || entry.FixedAsset ||
            (entry.originalEntry ? entry.originalEntry.FixedAsset : `Asset-${index + 1}`),
          originalEntry: entry.originalEntry || entry.OriginalEntry || {},
          errorCode: errorCode
        };
      });

      return {
        entries: processedEntries,
        summary: {
          totalCount: failedEntries.length,
          failedChanges: failedEntries.length,
          timestamp: new Date().toISOString()
        }
      };
    }

    /**
     * Format partial change results for display
     * @param {Array} successEntries - Array of successful entries
     * @param {Array} failedEntries - Array of failed entries  
     * @returns {Object} Formatted data for display
     */
    formatPartialChangeResultsForDisplay(successEntries, failedEntries) {
      console.log("ExportManager: Formatting partial change results for display...");

      const successData = this.formatChangeSuccessEntriesForDisplay(successEntries || []);
      const failureData = this.formatFailedChangeEntriesForDisplay(failedEntries || []);

      return {
        successEntries: successData.entries,
        failedEntries: failureData.entries,
        summary: {
          totalProcessed: (successEntries?.length || 0) + (failedEntries?.length || 0),
          successfulChanges: successEntries?.length || 0,
          failedChanges: failedEntries?.length || 0,
          successRate: ((successEntries?.length || 0) / Math.max(1, (successEntries?.length || 0) + (failedEntries?.length || 0))) * 100,
          timestamp: new Date().toISOString()
        }
      };
    }

    // =================
    // EXPORT METHODS
    // =================

    /**
     * Process API response data for export - handles the specific SAP API response formats
     * @param {Object} response - API response object (success or error)
     * @returns {Object} Processed data ready for export
     * @private
     */
    _processApiResponseForExport(response) {
      // Check if this is an error response
      if (response.error) {
        return {
          Status: "Error",
          ErrorCode: response.error.code || "Unknown",
          ErrorMessage: response.error.message || "Unknown error",
          Target: response.error.target || "",
          TransactionId: response.error.innererror?.ErrorDetails?.["@SAP__common.TransactionId"] || "",
          Timestamp: response.error.innererror?.ErrorDetails?.["@SAP__common.Timestamp"] || "",
          ComponentId: response.error.innererror?.ErrorDetails?.["@SAP__common.Application"]?.ComponentId || "",
          ServiceId: response.error.innererror?.ErrorDetails?.["@SAP__common.Application"]?.ServiceId || ""
        };
      }

      // Process success response (specific to SAP FixedAsset API)
      return {
        Status: "Success",
        CompanyCode: response.CompanyCode || "",
        MasterFixedAsset: response.MasterFixedAsset || "",
        FixedAsset: response.FixedAsset || "",
        AssetLifecycleStatus: response.AssetLifecycleStatus || "",
        FixedAssetDescription: response.FixedAssetDescription || "",
        AssetAdditionalDescription: response.AssetAdditionalDescription || "",
        AssetSerialNumber: response.AssetSerialNumber || "",
        InventoryNote: response.InventoryNote || "",
        InventoryIsCounted: response.InventoryIsCounted ? "Yes" : "No",
        CreatedByUser: response.CreatedByUser || "",
        CreationDate: response.CreationDate || "",
        LastChangedByUser: response.LastChangedByUser || "",
        LastChangeDate: response.LastChangeDate || "",
        LastChangeDateTime: response.LastChangeDateTime || ""
      };
    }

    /**
     * Export change results from batch processing
     * @param {Object} batchData - Batch processing results
     * @param {String} type - Type of export ("all", "success", "error")
     * @param {String} format - Export format ("xlsx", "csv", "pdf")
     * @returns {Promise}
     */
    exportChangeResults(batchData, type = "all", format = "xlsx") {
      console.log("ExportManager: Exporting change results...");
      return this.exportBatchResults(batchData, type, format);
    }

    /**
     * Export change validation errors
     * @param {Array} errors - Array of validation errors
     * @returns {Promise}
     */
    exportChangeValidationErrors(errors) {
      console.log("ExportManager: Exporting change validation errors...");

      if (!errors || errors.length === 0) {
        this._errorHandler.showWarning("No validation errors to export.");
        return Promise.reject(new Error("No validation errors to export."));
      }

      try {
        // Transform flat errors into grouped structure expected by existing method
        const groupedErrorsData = this._groupErrorsForExport(errors);
        return this.exportValidationErrorsToCSV(groupedErrorsData);
      } catch (error) {
        console.error("ExportManager: Error processing validation errors:", error);
        this._errorHandler.showError("Failed to process validation errors for export: " + error.message);
        return Promise.reject(error);
      }
    }

    /**
     * Exports validation errors to CSV file using CsvFormatter.
     * @param {Object} groupedErrorsData - Grouped errors data structure
     * @public
     */
    exportValidationErrorsToCSV(groupedErrorsData) {
      console.log("ExportManager: Exporting validation errors...");
      if (!groupedErrorsData || Object.keys(groupedErrorsData).length === 0) {
        this._errorHandler.showWarning("No errors available to export.");
        return Promise.reject(new Error("No errors available to export."));
      }

      try {
        // Convert grouped errors to flat array for CSV export
        const flatErrorsForCSV = [];
        Object.values(groupedErrorsData).forEach(({ category, errors }) => {
          errors.forEach((error) => {
            flatErrorsForCSV.push({
              SequenceNumber: error.sequenceId || "N/A",
              Asset: error.asset || "Unknown",
              Category: category,
              Field: error.field || "Unknown",
              Error: error.message || "Error Description Missing"
            });
          });
        });

        const fileName = `validation_errors_${new Date().toISOString().slice(0, 10)}.csv`;

        // Use CsvFormatter for consistent CSV generation
        return this._csvFormatter.format(flatErrorsForCSV)
          .then(blob => {
            triggerDownload(blob, fileName);
            this._errorHandler.showSuccess("Validation errors exported successfully.");
          })
          .catch(error => {
            console.error("ExportManager: Error exporting validation errors CSV:", error);
            this._errorHandler.showError("Failed to export validation errors.", error.message);
            throw error;
          });

      } catch (error) {
        console.error("ExportManager: Error preparing validation errors for CSV:", error);
        this._errorHandler.showError("Failed to prepare validation errors for export.", error.message);
        return Promise.reject(error);
      }
    }

    /**
     * Exports data to Excel, choosing the best method (SAP Spreadsheet or XLSX fallback).
     * @param {Array} data - Raw data to export.
     * @param {String} fileName - Desired file name (extension added automatically).
     * @param {Object} [options={}] - Export options.
     * @returns {Promise<void>}
     */
    exportToExcel(data, fileName, options = {}) {
      const reportType = options.reportType || "all";
      let processedData;

      try {
        // Delegate to DataTransformer to process rows for export
        processedData = this._dataTransformer.processRowsForExport(data, reportType);
        if (!processedData || processedData.length === 0) {
          const msg = `No data of type '${reportType}' available to export.`;
          this._errorHandler.showWarning(msg);
          return Promise.reject(new Error(msg));
        }
      } catch (error) {
        this._errorHandler.showError("Error processing data for Excel export: " + error.message);
        return Promise.reject(error);
      }

      const finalFileName = fileName.toLowerCase().endsWith(".xlsx") ? fileName : `${fileName}.xlsx`;
      const sheetName = `${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Records`;

      if (options.useSapSpreadsheet) {
        console.debug("ExportManager: Attempting SAP Spreadsheet export.");
        return this._sapSpreadsheetFormatter.formatAndDownload(processedData, finalFileName, {
          sheetName: sheetName,
          title: `Export Report - ${reportType}`
        })
          .then(() => {
            this._errorHandler.showSuccess("SAP Spreadsheet export started.");
          })
          .catch(sapError => {
            console.warn("ExportManager: SAP Spreadsheet failed, falling back to XLSX.", sapError);
            this._errorHandler.showWarning("SAP export failed, attempting fallback export.");

            // Fallback to XlsxFormatter
            return this._xlsxFormatter.formatSingleSheet(processedData, sheetName)
              .then(blob => {
                triggerDownload(blob, finalFileName);
                this._errorHandler.showSuccess("Fallback XLSX export successful.");
              })
              .catch(fallbackError => {
                this._errorHandler.showError("Fallback XLSX export also failed: " + fallbackError.message);
                throw fallbackError;
              });
          });
      } else {
        console.debug("ExportManager: Using XlsxFormatter export.");
        return this._xlsxFormatter.formatSingleSheet(processedData, sheetName)
          .then(blob => {
            triggerDownload(blob, finalFileName);
            this._errorHandler.showSuccess("XLSX export successful.");
          })
          .catch(error => {
            this._errorHandler.showError("XLSX export failed: " + error.message);
            throw error;
          });
      }
    }

    /**
     * Exports batch processing results in the specified format.
     * @param {Object} batchData - Raw batch results object.
     * @param {String} [type="all"] - Type of records to export ("all", "success", "error"/"errors").
     * @param {String} [format="xlsx"] - Desired format ("xlsx", "csv", "pdf").
     * @returns {Promise<void>}
     */
    exportBatchResults(batchData, type = "all", format = "xlsx") {
      try {
        console.log("ExportManager: Processing batch results with type:", type, "format:", format);

        // Process batch data using DataTransformer's implementation
        const batchResults = this._dataTransformer.processBatchResults(batchData, type);
        let processedData = batchResults.processedRecords || [];

        if (!processedData || processedData.length === 0) {
          const msg = `No batch data matching type '${type}' found to export.`;
          console.warn("ExportManager:", msg);
          this._errorHandler.showWarning(msg);
          return Promise.reject(new Error(msg));
        }

        const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
        const reportTypeDesc = type.toLowerCase().startsWith("error") ? "Error" :
          type.toLowerCase() === "success" ? "Success" : "All";
        const baseFileName = `Asset_Change_${reportTypeDesc}_Records_${timestamp}`;

        console.log("ExportManager: Starting export with format:", format, "filename:", baseFileName);

        // Use appropriate formatter based on format
        switch (format.toLowerCase()) {
          case "csv":
            const csvFileName = baseFileName + ".csv";
            console.log("ExportManager: Exporting to CSV:", csvFileName);

            return this._csvFormatter.format(processedData)
              .then(blob => {
                triggerDownload(blob, csvFileName);
                this._errorHandler.showSuccess("CSV export successful.");
              })
              .catch(error => {
                console.error("ExportManager: CSV export failed:", error);
                this._errorHandler.showError("CSV export failed: " + error.message);
                throw error;
              });

          case "pdf":
            const pdfFileName = baseFileName + ".pdf";
            const pdfOptions = {
              reportType: type,
              title: `Asset Change Results - ${reportTypeDesc}`,
              orientation: "landscape",
              pageSize: "A4"
            };
            console.log("ExportManager: Exporting to PDF:", pdfFileName, "with", processedData.length, "records");

            return this._pdfFormatter.format(processedData, pdfOptions)
              .then(blob => {
                triggerDownload(blob, pdfFileName);
                this._errorHandler.showSuccess("PDF export successful.");
              })
              .catch(error => {
                console.error("ExportManager: PDF export failed:", error);
                this._errorHandler.showError("PDF export failed: " + error.message);
                throw error;
              });

          case "xlsx":
          default:
            const xlsxFileName = baseFileName + ".xlsx";
            const sheetName = `${reportTypeDesc} Records`;
            console.log("ExportManager: Exporting to XLSX:", xlsxFileName, "with", processedData.length, "records");

            return this._sapSpreadsheetFormatter.formatAndDownload(processedData, xlsxFileName, {
              sheetName: sheetName,
              title: `Asset Change Report - ${reportTypeDesc}`
            })
              .catch(sapError => {
                console.warn("ExportManager: SAP Spreadsheet failed for XLSX batch export, falling back to basic XLSX.", sapError);
                return this._xlsxFormatter.formatSingleSheet(processedData, sheetName);
              })
              .then(blob => {
                if (blob) {
                  triggerDownload(blob, xlsxFileName);
                }
                this._errorHandler.showSuccess("XLSX export successful.");
              })
              .catch(error => {
                console.error("ExportManager: XLSX export failed:", error);
                this._errorHandler.showError("XLSX export failed: " + error.message);
                throw error;
              });
        }
      } catch (error) {
        console.error("ExportManager: Error during batch export:", error);
        this._errorHandler.showError("Export failed: " + error.message);
        return Promise.reject(error);
      }
    }

    /**
     * Downloads the template file using TemplateGenerator.
     * @returns {Promise<void>}
     */
    downloadChangeTemplate() {
      return this._templateGenerator.generateChangeTemplateBlob()
        .then(blob => {
          const fileName = this._templateGenerator.getChangeTemplateFileName();
          triggerDownload(blob, fileName);
          this._errorHandler.showSuccess("Template downloaded successfully.");
        })
        .catch(error => {
          this._errorHandler.showError("Template download failed: " + error.message);
          throw error;
        });
    }

    // =================
    // UTILITY METHODS
    // =================

    /**
     * Group flat errors array into the structure expected by exportValidationErrorsToCSV
     * @param {Array} errors - Flat array of error objects
     * @returns {Object} Grouped errors data
     * @private
     */
    _groupErrorsForExport(errors) {
      const grouped = {
        "Validation Errors": {
          category: "Validation Errors",
          errors: errors.map(error => ({
            sequenceId: error.sequenceId || error.SequenceId || "N/A",
            field: error.field || error.Field || "Unknown",
            message: error.message || error.Message || "Error description missing",
            asset: error.asset || error.Asset || "Unknown"
          }))
        }
      };

      return grouped;
    }

    /**
     * Gets the data transformer instance for use by other utilities.
     * @returns {DataTransformer} The data transformer instance
     * @public
     */
    getDataTransformer() {
      return this._dataTransformer;
    }

    /**
     * Gets the error handler instance for use by other utilities.
     * @returns {ErrorHandler} The error handler instance
     * @public
     */
    getErrorHandler() {
      return this._errorHandler;
    }
  };
});