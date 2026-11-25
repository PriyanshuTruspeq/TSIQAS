sap.ui.define([
    "sap/ui/export/Spreadsheet",
    "sap/ui/export/library"
], function (Spreadsheet, exportLibrary) {
    "use strict";

    const EdmType = exportLibrary.EdmType;

    return class SapSpreadsheetFormatter {

        getFileExtension() { return ".xlsx"; }
        getMimeType() { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; }

        /**
         * Exports data using sap.ui.export.Spreadsheet.
         * This directly triggers the download via the library.
         * @param {Array<object>} data - Processed data array.
         * @param {string} fileName - Desired file name.
         * @param {object} [options={}] - Options like context metadata.
         * @param {string} [options.sheetName="Asset Change Report"] - Name for the sheet.
         * @param {string} [options.title="Asset Master Change Report"] - Title metadata.
         * @param {string} [options.application="Asset Master Change"] - Application metadata.
         * @returns {Promise<void>} Promise resolving when build starts, or rejecting on error.
         */
        formatAndDownload(data, fileName, options = {}) {
            return new Promise((resolve, reject) => {
                try {
                    if (!data || data.length === 0) {
                        reject(new Error("No data provided for SAP Spreadsheet export."));
                        return;
                    }

                    // Sort data by sequence number before processing
                    data.sort((a, b) => {
                        const seqA = parseInt(a.SequenceNumber) || 0;
                        const seqB = parseInt(b.SequenceNumber) || 0;
                        return seqA - seqB;
                    });

                    // Define specific columns for Asset Change data with proper formatting
                    const columns = [
                        { label: "Sequence Number", property: "SequenceNumber", type: EdmType.Number, width: 15 },
                        { label: "Company Code", property: "CompanyCode", type: EdmType.String, width: 15 },
                        { label: "Fixed Asset", property: "FixedAsset", type: EdmType.String, width: 15 },
                        { label: "Status", property: "Status", type: EdmType.String, width: 12 },
                        { label: "Asset Description", property: "AssetDescription", type: EdmType.String, width: 30 },
                        { label: "Serial Number", property: "SerialNumber", type: EdmType.String, width: 20 },
                        { label: "Inventory Note", property: "InventoryNote", type: EdmType.String, width: 25 },
                        { label: "Total Changes", property: "TotalChanges", type: EdmType.Number, width: 15 },
                        { label: "Changed Fields", property: "ChangedFieldsSummary", type: EdmType.String, width: 50 },
                        { label: "Has Blank Assignments", property: "HasBlankAssignments", type: EdmType.String, width: 20 },
                        { label: "Message", property: "MessageText", type: EdmType.String, width: 40 },
                        { label: "Error Details", property: "ErrorDetails", type: EdmType.String, width: 40 },
                        { label: "Export Timestamp", property: "ExportTimestamp", type: EdmType.DateTimeOffset, width: 20 }
                    ];

                    // Transform data to match column structure and improve readability
                    const transformedData = data.map(item => ({
                        SequenceNumber: parseInt(item.SequenceNumber) || 0,
                        CompanyCode: item.CompanyCode || "",
                        FixedAsset: item.FixedAsset || "",
                        Status: item.Status || "",
                        AssetDescription: item.AssetDescription || "",
                        SerialNumber: item.SerialNumber || "",
                        InventoryNote: item.InventoryNote || "",
                        TotalChanges: (item.ChangesSummary && item.ChangesSummary.totalChanges) || 0,
                        ChangedFieldsSummary: item.ChangedFieldsSummary || "",
                        HasBlankAssignments: item.HasBlankAssignments ? "Yes" : "No",
                        MessageText: item.Status === "Error" ? (item.ErrorMessage || "") : (item.Message || ""),
                        ErrorDetails: item.ErrorDetails || "",
                        ExportTimestamp: item.ExportTimestamp || new Date().toISOString()
                    }));

                    const exportSettings = {
                        workbook: {
                            columns: columns,
                            context: {
                                application: options.application || "Asset Master Change",
                                sheetName: options.sheetName || "Asset Change Report",
                                title: options.title || "Asset Master Change Report",
                                modifiedBy: "System User",
                                metaSheetName: "Metadata"
                            },
                            hierarchyLevel: "Level"
                        },
                        dataSource: transformedData,
                        fileName: fileName.toLowerCase().endsWith(".xlsx") ? fileName : `${fileName}.xlsx`,
                        worker: false // Better performance for moderate datasets
                    };

                    const oSpreadsheet = new Spreadsheet(exportSettings);
                    oSpreadsheet.build()
                        .then(() => {
                            console.log("SAP Spreadsheet build completed successfully.");
                            resolve(); // Resolve when build is done (download is triggered)
                        })
                        .catch((error) => {
                            console.error("SAP Spreadsheet build error:", error);
                            reject(error); // Reject the promise on build error
                        })
                        .finally(() => {
                            if (oSpreadsheet && oSpreadsheet.destroy) {
                                oSpreadsheet.destroy(); // Clean up
                            }
                        });

                } catch (error) {
                    console.error("Error setting up SAP Spreadsheet export:", error);
                    reject(error);
                }
            });
        }
    };
});