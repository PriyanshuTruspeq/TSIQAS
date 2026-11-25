sap.ui.define([], function () {
    "use strict";

    // Helper to ensure XLSX utils are loaded for CSV conversion
    const checkXLSXUtils = () => {
        if (typeof window.XLSX === 'undefined' || typeof window.XLSX.utils === 'undefined') {
            throw new Error("XLSX library utils not loaded. CSV export requires them for robust conversion.");
        }
    };

    return class CsvFormatter {
        getFileExtension() { return ".csv"; }
        getMimeType() { return "text/csv;charset=utf-8;"; }

        /**
         * Creates a CSV Blob from an array of data.
         * Uses XLSX utils for reliable CSV conversion with improved formatting.
         * @param {Array<object>} data - Array of processed data objects.
         * @param {object} [options={}] - Export options.
         * @returns {Promise<Blob>} Promise resolving with the CSV Blob.
         */
        format(data, options = {}) {
            return new Promise((resolve, reject) => {
                try {
                    checkXLSXUtils();

                    if (!data || data.length === 0) {
                        reject(new Error("No data provided for CSV export."));
                        return;
                    }

                    // Sort data by sequence number before processing
                    data.sort((a, b) => {
                        const seqA = parseInt(a.SequenceNumber) || 0;
                        const seqB = parseInt(b.SequenceNumber) || 0;
                        return seqA - seqB;
                    });

                    // Transform data for better CSV format with readable column names
                    const transformedData = data.map(item => ({
                        "Sequence Number": parseInt(item.SequenceNumber) || 0,
                        "Company Code": item.CompanyCode || "",
                        "Fixed Asset": item.FixedAsset || "",
                        "Status": item.Status || "",
                        "Asset Description": item.AssetDescription || "",
                        "Serial Number": item.SerialNumber || "",
                        "Inventory Note": item.InventoryNote || "",
                        "Total Changes": (item.ChangesSummary && item.ChangesSummary.totalChanges) || 0,
                        "Create Count": (item.ChangesSummary && item.ChangesSummary.createCount) || 0,
                        "Change Count": (item.ChangesSummary && item.ChangesSummary.changeCount) || 0,
                        "Delete Count": (item.ChangesSummary && item.ChangesSummary.deleteCount) || 0,
                        "Changed Fields": item.ChangedFieldsSummary || "",
                        "Has Blank Assignments": item.HasBlankAssignments ? "Yes" : "No",
                        "Message": item.Status === "Error" ? (item.ErrorMessage || "") : (item.Message || ""),
                        "Error Details": item.ErrorDetails || "",
                        "Error Code": item.ErrorCode || "",
                        "Export Timestamp": item.ExportTimestamp || new Date().toISOString(),
                        "Record Type": item.RecordType || "all"
                    }));

                    // Create worksheet with proper formatting
                    const worksheet = window.XLSX.utils.json_to_sheet(transformedData);
                    
                    // Set column widths for better readability
                    const columnWidths = [
                        { wch: 15 }, // Sequence Number
                        { wch: 12 }, // Company Code
                        { wch: 15 }, // Fixed Asset
                        { wch: 10 }, // Status
                        { wch: 30 }, // Asset Description
                        { wch: 20 }, // Serial Number
                        { wch: 25 }, // Inventory Note
                        { wch: 12 }, // Total Changes
                        { wch: 12 }, // Create Count
                        { wch: 12 }, // Change Count
                        { wch: 12 }, // Delete Count
                        { wch: 50 }, // Changed Fields
                        { wch: 20 }, // Has Blank Assignments
                        { wch: 40 }, // Message
                        { wch: 40 }, // Error Details
                        { wch: 15 }, // Error Code
                        { wch: 20 }, // Export Timestamp
                        { wch: 12 }  // Record Type
                    ];
                    
                    worksheet['!cols'] = columnWidths;

                    // Convert to CSV with proper formatting
                    const csvContent = window.XLSX.utils.sheet_to_csv(worksheet, {
                        FS: ',',        // Field separator
                        RS: '\r\n',     // Record separator
                        strip: false,   // Don't strip whitespace
                        blankrows: false // Skip blank rows
                    });

                    // Add BOM for UTF-8 compatibility in Excel
                    const blob = new Blob(["\uFEFF" + csvContent], { 
                        type: this.getMimeType() 
                    });
                    
                    resolve(blob);
                } catch (error) {
                    console.error("Error formatting CSV:", error);
                    reject(error);
                }
            });
        }

        /**
         * Helper method to escape CSV values properly
         * @private
         */
        _escapeCsvValue(value) {
            if (value === null || value === undefined) return "";
            
            const stringValue = String(value);
            
            // If value contains comma, quote, or newline, wrap in quotes and escape quotes
            if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
                return '"' + stringValue.replace(/"/g, '""') + '"';
            }
            
            return stringValue;
        }
    };
});