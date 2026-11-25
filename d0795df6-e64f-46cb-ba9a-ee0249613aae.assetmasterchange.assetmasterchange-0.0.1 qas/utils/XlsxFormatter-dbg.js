sap.ui.define([], function () {
    "use strict";

    return class XlsxFormatter {

        constructor() {
            this._validateXLSX();
        }

        /**
         * Validates that XLSX library is available
         * @private
         */
        _validateXLSX() {
            if (!window.XLSX) {
                console.error("XlsxFormatter: XLSX library not loaded");
                throw new Error("XLSX library is required but not loaded");
            }
        }

        /**
         * Creates a worksheet from headers and data
         * @param {Array<string>} headers - Column headers
         * @param {Array<Object>} data - Data rows
         * @returns {Promise<Object>} Worksheet object
         */
        async createWorksheet(headers, data = []) {
            try {
                // Create array of arrays format
                const aoa = [headers];

                // Add data rows
                data.forEach(row => {
                    const rowData = headers.map(header => {
                        const value = row[header];
                        // Handle null/undefined
                        if (value === null || value === undefined) {
                            return "";
                        }
                        // Handle dates
                        if (value instanceof Date) {
                            return value.toISOString().split('T')[0];
                        }
                        // Handle booleans
                        if (typeof value === 'boolean') {
                            return value ? "TRUE" : "FALSE";
                        }
                        // Convert to string
                        return String(value);
                    });
                    aoa.push(rowData);
                });

                // Create worksheet
                const ws = window.XLSX.utils.aoa_to_sheet(aoa);

                // Apply header styles (basic)
                this._styleHeaders(ws, headers.length);

                return ws;
            } catch (error) {
                console.error("XlsxFormatter: Error creating worksheet", error);
                throw error;
            }
        }

        /**
         * Creates a workbook for the change template
         * @param {Array<string>} headers - Column headers
         * @param {Object} exampleRow - Example data row
         * @returns {Promise<Object>} Workbook object
         */
        async createWorkbookForChangeTemplate(headers, exampleRow) {
            try {
                const wb = window.XLSX.utils.book_new();

                // Create main data sheet
                const ws = await this.createWorksheet(headers, [exampleRow]);

                // Set column widths
                ws['!cols'] = this._calculateColumnWidths(headers);

                // Add to workbook
                window.XLSX.utils.book_append_sheet(wb, ws, "Change_Data");

                return wb;
            } catch (error) {
                console.error("XlsxFormatter: Error creating workbook", error);
                throw error;
            }
        }

        /**
         * Converts workbook to blob
         * @param {Object} workbook - XLSX workbook object
         * @returns {Promise<Blob>} Blob object
         */
        async workbookToBlob(workbook) {
            try {
                // Write workbook to binary string
                const wbout = window.XLSX.write(workbook, {
                    bookType: 'xlsx',
                    type: 'array'
                });

                // Create blob
                const blob = new Blob([wbout], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                });

                return blob;
            } catch (error) {
                console.error("XlsxFormatter: Error converting workbook to blob", error);
                throw error;
            }
        }

        /**
         * Exports data to Excel file
         * @param {Array<Object>} data - Data to export
         * @param {string} filename - Output filename
         * @param {Object} options - Export options
         */
        exportToExcel(data, filename, options = {}) {
            try {
                if (!data || !Array.isArray(data) || data.length === 0) {
                    throw new Error("No data to export");
                }

                // Get headers from first row
                const headers = Object.keys(data[0]);

                // Create worksheet
                const ws = window.XLSX.utils.json_to_sheet(data, {
                    header: headers,
                    skipHeader: false
                });

                // Apply column widths
                if (options.columnWidths) {
                    ws['!cols'] = options.columnWidths;
                } else {
                    ws['!cols'] = this._calculateColumnWidths(headers, data);
                }

                // Create workbook
                const wb = window.XLSX.utils.book_new();
                window.XLSX.utils.book_append_sheet(wb, ws, options.sheetName || "Data");

                // Write file
                window.XLSX.writeFile(wb, filename);

                return true;
            } catch (error) {
                console.error("XlsxFormatter: Error exporting to Excel", error);
                throw error;
            }
        }

        /**
         * Exports multiple sheets to Excel file
         * @param {Object} sheets - Object with sheet names as keys and data arrays as values
         * @param {string} filename - Output filename
         */
        exportMultipleSheets(sheets, filename) {
            try {
                const wb = window.XLSX.utils.book_new();

                Object.entries(sheets).forEach(([sheetName, data]) => {
                    if (data && Array.isArray(data) && data.length > 0) {
                        const ws = window.XLSX.utils.json_to_sheet(data);

                        // Auto-size columns
                        const headers = Object.keys(data[0]);
                        ws['!cols'] = this._calculateColumnWidths(headers, data);

                        window.XLSX.utils.book_append_sheet(wb, ws, sheetName);
                    }
                });

                window.XLSX.writeFile(wb, filename);
                return true;
            } catch (error) {
                console.error("XlsxFormatter: Error exporting multiple sheets", error);
                throw error;
            }
        }

        /**
         * Calculate column widths based on content
         * @private
         * @param {Array<string>} headers - Column headers
         * @param {Array<Object>} data - Data rows (optional)
         * @returns {Array<Object>} Column width configuration
         */
        _calculateColumnWidths(headers, data) {
            const widths = headers.map(header => {
                let maxLength = header.length;

                // Check data rows if provided
                if (data && data.length > 0) {
                    data.forEach(row => {
                        const value = String(row[header] || '');
                        maxLength = Math.max(maxLength, value.length);
                    });
                }

                // Set width with min/max bounds
                const width = Math.min(Math.max(maxLength + 2, 10), 50);
                return { wch: width };
            });

            return widths;
        }

        /**
         * Apply basic styling to headers
         * @private
         * @param {Object} ws - Worksheet object
         * @param {number} headerCount - Number of headers
         */
        _styleHeaders(ws, headerCount) {
            // Note: XLSX-style library would be needed for full styling
            // This is a placeholder for basic structure

            // Add autofilter to headers
            const range = window.XLSX.utils.decode_range(ws['!ref']);
            ws['!autofilter'] = {
                ref: window.XLSX.utils.encode_range({
                    s: { r: 0, c: 0 },
                    e: { r: range.e.r, c: headerCount - 1 }
                })
            };
        }

        /**
         * Format date for Excel
         * @param {Date|string} date - Date to format
         * @returns {string} Formatted date string
         */
        formatDateForExcel(date) {
            if (!date) return "";

            const d = date instanceof Date ? date : new Date(date);
            if (isNaN(d.getTime())) return "";

            return d.toISOString().split('T')[0];
        }

        /**
         * Clean data for Excel export
         * @param {*} value - Value to clean
         * @returns {*} Cleaned value
         */
        cleanValueForExcel(value) {
            if (value === null || value === undefined) {
                return "";
            }

            if (typeof value === 'boolean') {
                return value ? "Yes" : "No";
            }

            if (value instanceof Date) {
                return this.formatDateForExcel(value);
            }

            if (typeof value === 'object') {
                return JSON.stringify(value);
            }

            // Remove special characters that might cause issues
            return String(value)
                .replace(/[\r\n\t]/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
        }

        /**
         * Parse CSV content using XLSX
         * @param {string} csvContent - CSV content as string
         * @returns {Array<Object>} Parsed data
         */
        parseCSV(csvContent) {
            try {
                const wb = window.XLSX.read(csvContent, { type: 'string' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                return window.XLSX.utils.sheet_to_json(ws);
            } catch (error) {
                console.error("XlsxFormatter: Error parsing CSV", error);
                throw error;
            }
        }
    };
});