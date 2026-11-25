sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast"
], function (Controller, Filter, FilterOperator, JSONModel, MessageToast) {
    "use strict";

    return Controller.extend("assetmasterchange.controller.BaseFilterController", {

        /**
         * Initialize filter model for asset master changes
         */
        onInitFilters: function () {
            if (!this.getView().getModel("filterModel")) {
                const oFilterModel = new JSONModel({
                    // Asset identification filters
                    companyCode: "",
                    assetClass: "",
                    description: "",

                    // Asset location and assignment
                    wbsElement: "",
                    plant: "",
                    costCenter: "",

                    // Asset attributes
                    serialNumber: "",
                    depreciationKey: "",
                    assetBlock: "",

                    // Status and change filters
                    status: "all", // all, valid, invalid
                    changes: "", // filter for assets with specific changes

                    // Date filters for acquisition
                    acquisitionDateFrom: null,
                    acquisitionDateTo: null,
                    capitalizationYear: ""
                });
                this.getView().setModel(oFilterModel, "filterModel");
            }

            // Set default date range to current year
            this._initializeDefaultDateRange();
        },

        /**
         * Initialize default date range to current year
         * @private
         */
        _initializeDefaultDateRange: function () {
            const oFilterModel = this.getView().getModel("filterModel");
            if (!oFilterModel.getProperty("/acquisitionDateFrom")) {
                const today = new Date();
                const firstDayOfYear = new Date(today.getFullYear(), 0, 1);
                oFilterModel.setProperty("/acquisitionDateFrom", firstDayOfYear);
                oFilterModel.setProperty("/acquisitionDateTo", today);
            }
        },

        /**
         * Apply all filters to the asset master entries table
         */
        applyFilters: function () {
            const oTable = this.byId("assetTable");
            const oBinding = oTable?.getBinding("items");
            const oFilterModel = this.getView().getModel("filterModel");

            if (!oBinding) {
                console.error("Cannot find table binding for assetTable");
                return;
            }

            if (!oFilterModel) {
                console.error("Cannot find filter model");
                return;
            }

            const aFilters = this._buildFilters(oFilterModel);

            // Apply filters to table
            if (aFilters.length > 0) {
                const oFilter = new Filter({
                    filters: aFilters,
                    and: true
                });
                oBinding.filter(oFilter);
            } else {
                oBinding.filter([]);
            }

            // Update filtered count and show message
            this._updateFilteredCount(oBinding);
        },

        /**
         * Build filters array from filter model
         * @private
         * @param {sap.ui.model.json.JSONModel} oFilterModel - Filter model
         * @returns {Array} Array of filters
         */
        _buildFilters: function (oFilterModel) {
            const aFilters = [];

            // Company Code filter
            const companyCode = oFilterModel.getProperty("/companyCode");
            if (companyCode) {
                aFilters.push(new Filter("CompanyCode", FilterOperator.Contains, companyCode));
            }

            // Asset Class filter
            const assetClass = oFilterModel.getProperty("/assetClass");
            if (assetClass) {
                aFilters.push(new Filter("AssetClass", FilterOperator.Contains, assetClass));
            }

            // Description filter
            const description = oFilterModel.getProperty("/description");
            if (description) {
                aFilters.push(new Filter("FixedAssetDescription", FilterOperator.Contains, description));
            }

            // WBS Element filter
            const wbsElement = oFilterModel.getProperty("/wbsElement");
            if (wbsElement) {
                aFilters.push(new Filter("YY1_WBS_ELEMENT", FilterOperator.Contains, wbsElement));
            }

            // Plant filter
            const plant = oFilterModel.getProperty("/plant");
            if (plant) {
                aFilters.push(new Filter("Plant", FilterOperator.Contains, plant));
            }

            // Cost Center filter
            const costCenter = oFilterModel.getProperty("/costCenter");
            if (costCenter) {
                aFilters.push(new Filter("CostCenter", FilterOperator.Contains, costCenter));
            }

            // FIXED: Status filter for validation status
            const status = oFilterModel.getProperty("/status");
            if (status && status !== "all") {
                console.log("Applying status filter:", status);
                
                // Create multiple filters to handle different possible status values
                const statusFilters = [];
                
                if (status === "valid") {
                    // Check for both "Valid" and "valid" (case insensitive)
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "Valid"));
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "valid"));
                } else if (status === "invalid") {
                    // Check for both "Invalid" and "invalid" (case insensitive)
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "Invalid"));
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "invalid"));
                    // Also include entries with validation errors
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "Error"));
                    statusFilters.push(new Filter("Status", FilterOperator.EQ, "error"));
                }
                
                if (statusFilters.length > 0) {
                    // Use OR logic for the status variations
                    aFilters.push(new Filter({
                        filters: statusFilters,
                        and: false // OR logic
                    }));
                }
            }

            // Changes filter
            const changes = oFilterModel.getProperty("/changes");
            if (changes) {
                aFilters.push(new Filter("HasChanges", FilterOperator.EQ, changes === "yes"));
            }

            console.log("Total filters applied:", aFilters.length);
            return aFilters;
        },

        /**
         * Update filtered count in the model and show message
         * @private
         * @param {sap.ui.model.ListBinding} oBinding - Table binding
         */
        _updateFilteredCount: function (oBinding) {
            const iFilteredCount = oBinding.getLength();
            const oAssetModel = this.getView().getModel("assetMasterEntries");

            if (!oAssetModel) {
                console.error("assetMasterEntries model not found");
                return;
            }

            // Update filtered count
            oAssetModel.setProperty("/filteredCount", iFilteredCount);

            // Show filter result message
            const totalEntries = oAssetModel.getProperty("/entries")?.length || 0;
            const filterModel = this.getView().getModel("filterModel");
            const currentStatus = filterModel?.getProperty("/status") || "all";
            
            if (currentStatus !== "all") {
                MessageToast.show(`Showing ${iFilteredCount} ${currentStatus} entries of ${totalEntries} total`);
            } else {
                MessageToast.show(`Showing ${iFilteredCount} of ${totalEntries} asset entries`);
            }
        },

        /**
         * Handle filter input changes
         */
        onFilterChange: function () {
            this.applyFilters();
        },

        /**
         * Handle status filter ComboBox changes
         * @param {sap.ui.base.Event} oEvent - Selection change event
         */
        onStatusFilterChange: function (oEvent) {
            const selectedItem = oEvent.getParameter("selectedItem");
            if (selectedItem) {
                const selectedKey = selectedItem.getKey();
                console.log("BaseFilterController: Status filter changed to", selectedKey);

                // Update filter model
                this.getView().getModel("filterModel").setProperty("/status", selectedKey);
                
                // Apply filters immediately
                this.applyFilters();
                
                // Debug: Log current data to verify status values
                this._debugStatusValues();
            }
        },

        /**
         * Debug method to check actual status values in the data
         * @private
         */
        _debugStatusValues: function () {
            const assetModel = this.getView().getModel("assetMasterEntries");
            const entries = assetModel?.getProperty("/entries") || [];
            
            console.log("Debug: Total entries:", entries.length);
            
            const statusCounts = {};
            entries.forEach(entry => {
                const status = entry.Status;
                statusCounts[status] = (statusCounts[status] || 0) + 1;
            });
            
            console.log("Debug: Status distribution:", statusCounts);
            
            // Show first few entries for debugging
            console.log("Debug: First 3 entries status:", 
                entries.slice(0, 3).map(e => ({ 
                    AssetNumber: e.FixedAsset || e.AssetNumber, 
                    Status: e.Status 
                }))
            );
        },

        /**
         * Handle changes filter changes
         * @param {sap.ui.base.Event} oEvent - Selection change event
         */
        onChangesFilterChange: function (oEvent) {
            const selectedItem = oEvent.getParameter("selectedItem");
            if (selectedItem) {
                const selectedKey = selectedItem.getKey();
                this.getView().getModel("filterModel").setProperty("/changes", selectedKey);
                this.applyFilters();
            }
        },

        /**
         * Reset all filters to default values
         */
        onResetFilters: function () {
            const oFilterModel = this.getView().getModel("filterModel");

            // Reset all filter properties
            oFilterModel.setData({
                companyCode: "",
                assetClass: "",
                description: "",
                wbsElement: "",
                plant: "",
                costCenter: "",
                serialNumber: "",
                depreciationKey: "",
                assetBlock: "",
                status: "all",
                changes: "",
                acquisitionDateFrom: null,
                acquisitionDateTo: null,
                capitalizationYear: ""
            });

            // Reset date range
            this._initializeDefaultDateRange();

            // Clear table filters
            const oTable = this.getView().byId("assetTable");
            const oBinding = oTable?.getBinding("items");

            if (oBinding) {
                oBinding.filter([]);
                this._updateFilteredCount(oBinding);
            }

            MessageToast.show("All filters have been reset");
        },

        /**
         * Refresh table data and debug status values
         */
        onRefreshTable: function () {
            const assetTable = this.byId("assetTable");
            if (assetTable) {
                const binding = assetTable.getBinding("items");
                if (binding) {
                    binding.refresh();
                }
            }
            
            // Debug current status values after refresh
            this._debugStatusValues();
            
            MessageToast.show("Asset table refreshed");
        }
    });
});