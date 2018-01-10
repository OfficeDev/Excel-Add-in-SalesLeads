/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


(function () {
    "use strict";

    // Create a namespace to hold application-wide settings with primitive data types.
    let SalesLeadApp = window.SalesLeadApp || {};

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            const element = document.querySelector('.ms-MessageBanner');
            FabricComponents.messageBanner = new FabricComponents.MessageBanner(element);
            FabricComponents.messageBanner.hideBanner();

            $('#import-data-button').click(SalesLeadApp.getData);
            $('#analyze-button').click(SalesLeadApp.analyse);
            $('#get-current-file-button').click(SalesLeadApp.showFileAsByteArray);
            
            // Verify the Office host supports 1.7 and earlier Excel APIs.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.6)) {
                SalesLeadApp.showNotification('Sorry. The Sales Lead add-in uses features that are not available in your version of Office.');
                document.getElementById('import-data-button').disabled = true;
            }           
        });
    };

    // The row Range addresses in the CustomersTable that
    // can be used as the targets of hyperlinks in the salesperson's
    // worksheet.
    SalesLeadApp.customerRowAddresses = [];

    // The open sales opportunities for a specific salesperson as a
    // an array of values to populate the table on the salesperson's 
    // sheet.
    SalesLeadApp.salesPersonLeads = []

    // The closed sales from last year for a specific salesperson as a 
    // RangeView object.
    SalesLeadApp.filteredClosedSales = null;

    // The closed sales as an array of values that can used to populate
    // a chart.
    SalesLeadApp.closedSales = [];

    // The salesperson to analyze
    SalesLeadApp.selectedSalesPerson = '';

    // Gets sales data, and create data sheets and tables.
    SalesLeadApp.getData = function (event) {

        const salesResponse = $.ajax({
            url: 'MockDynamicsData/SalesLeadsData.json',
            dataType: 'json',
            cache: false
        });

        const customerResponse = $.ajax({
            url: 'MockDynamicsData/Customers.json',
            dataType: 'json',
            cache: false
        });

        $.when(salesResponse, customerResponse)
        .then(function (resolvedSalesLeadData, resolvedCustomerData) {

            // The sales lead and customer data is converted from JSON
            // to arrays that can populate Ranges.
            const salesLeadData = resolvedSalesLeadData[0];
            const customerData = resolvedCustomerData[0];

            // Return them as a single object so they can be passed to
            // the callback of the next "then".
            return { salesLeadData, customerData };
        })
        .then(function ({ salesLeadData, customerData }) {
            return Excel.run(function (context) {
                const sheets = context.workbook.worksheets;
                createCustomersSheetAndTable(sheets, customerData);
                createOpportunitySheetAndTable(sheets, salesLeadData);
                document.getElementById('analyze-button').disabled = false;
                sheets.getItem('Sheet1').delete();

                return context.sync()         
            })               
            .catch(SalesLeadApp.errorHandler);
         })

        // Queue commands to create the Customer data worksheet and table.
        function createCustomersSheetAndTable(sheets, data) {
            const customerSheet = sheets.add();
            customerSheet.name = 'Customers';
            customerSheet.tabColor = 'red';

            // Freeze top (header) row so it will stay visible when user scrolls down.
            customerSheet.freezePanes.freezeRows(1);
            customerSheet.activate();

            let customersTable = customerSheet.tables.add('A1:E1', true);
            customersTable.name = 'CustomersTable';
            customersTable.getHeaderRowRange().values = [[
                'Name',
                'Phone',
                'City',
                'Primary Contact',
                'Email'
            ]];

            const tableBodyData = data.map(customer => [
                customer.Name,
                customer.Phone,
                customer.City,
                customer.PrimaryContact,
                customer.Email
            ]);
            customersTable.rows.add(null, tableBodyData);

            customerSheet.getUsedRange().format.autofitColumns();
            customerSheet.getUsedRange().format.autofitRows();

            // Save customer row addresses to use in hyperlinks in
            // another method. Do it now to save a context.sync.
            data.forEach(function (customer, index) {
                SalesLeadApp.customerRowAddresses.push({
                    'Name': customer.Name,

                    // "index+2" because arrays are 0-based, but Range addresses 
                    // are 1-based and there is a header row. 
                    'Address': `Customers!A${index + 2}:E${index + 2}`
                });
            });
        }


        // Queue commands to create the Opportunities worksheet and 
        // data table.
        function createOpportunitySheetAndTable(sheets, data) {
            const opportunitySheet = sheets.add();
            opportunitySheet.name = 'Opportunities';
            opportunitySheet.tabColor = 'red';

            // Freeze top (header) row so it remains visible when user scrolls down.
            opportunitySheet.freezePanes.freezeRows(1);
            opportunitySheet.activate();

            let salesLeadTable = opportunitySheet.tables.add('A1:F1', true);
            salesLeadTable.name = 'SalesLeadTable';
            salesLeadTable.getHeaderRowRange().values = [[
                'Owner',
                'Account',
                'Topic',
                'Probability',
                'Est. Close Date',
                'Est. Revenue'
            ]];
            salesLeadTable.columns.getItemAt(5).getRange().numberFormat = [['$#,##0.00']];
         //   salesLeadTable.columns.getItemAt(7).getRange().numberFormat = [['$#,##0.00']];

            const tableBodyData = data.map(lead => [
                lead.Owner,
                lead.Account,
                lead.Topic,
                lead.Probability,
                lead.EstCloseDate,
                lead.EstRevenue
            ]);
            salesLeadTable.rows.add(null, tableBodyData);

            opportunitySheet.getUsedRange().format.autofitColumns();
            opportunitySheet.getUsedRange().format.autofitRows();
        }
    }
  
    // Create a workwheet with a specific salesperson's unclosed sales
    // leads. Use conditional formatting to highlight the relative value of 
    // the leads. Create a hidden worksheet to hold the salesperson's sales last 
    // year. Create a chart on the salesperson's worksheet showing last year's sales.
    SalesLeadApp.analyse = function (event) {
        return Excel.run(function (context) {

            const selectedSalesPersonRange = context.workbook.getSelectedRange().load('values');
            const saleLeadRows = context.workbook.tables.getItem('SalesLeadTable')
                .rows.load('count, values');

            return context.sync()
                .then(function () {
                    SalesLeadApp.selectedSalesPerson = selectedSalesPersonRange.values[0][0];

                    for (let i = 0; i < saleLeadRows.count; i++) {
                        let row = saleLeadRows.items[i];
                        if (
                            (row.values[0][0] === SalesLeadApp.selectedSalesPerson)
                            // Excel stores dates internally as days since 1899-12-31.
                            // 42986 is 2017-9-10.
                            && (row.values[0][4] > 42986)
                        ) {
                            SalesLeadApp.salesPersonLeads.push({
                                'Account': row.values[0][1],
                                'Probability': row.values[0][3],
                                'EstRevenue': row.values[0][5]
                            });
                        }
                    }

                    const sheets = context.workbook.worksheets;
                    SalesLeadApp.filteredClosedSales = getFilteredData(sheets);
                })
                .then(context.sync)
                .then(function () {
                    cacheFilteredData(SalesLeadApp.filteredClosedSales);
                    const sheets = context.workbook.worksheets;
                    createTempSheet(sheets);
                    createAnalysisSheet(sheets, SalesLeadApp.salesPersonLeads);

                    // Empty the salesPersonLeads so that the analyze() method can be run more than once.
                    SalesLeadApp.salesPersonLeads = [];
                    document.getElementById('get-current-file-button').disabled = false;
                })
                .then(context.sync)
        })
        .catch(SalesLeadApp.errorHandler);

        function getFilteredData(sheets) {

            // We need a temporary filtering of the opportunity data to create
            // the chart on the salesperson's personal worksheet.
            const salesLeadTable = sheets.getItem('Opportunities').tables.getItem('SalesLeadTable');
            const filteredClosedSales = filterSalesLeads(sheets, salesLeadTable);

            // Clear the filters so all the data appears on the Opportunity worksheet.
            salesLeadTable.clearFilters();
            return filteredClosedSales;

            // Queue commands to filter in only the sales leads owned by specific user 
            // and that closed last year. This data is used to make a range on a 
            // hidden sheet that will be the basis of the chart on the salesperson's
            // worksheet.
            function filterSalesLeads(sheets, tableToFilter) {
                const ownerFilter = tableToFilter.columns.getItem('Owner').filter;
                ownerFilter.applyValuesFilter([SalesLeadApp.selectedSalesPerson]);
                const dateFilter = tableToFilter.columns.getItem('Est. Close Date').filter;
                dateFilter.applyDynamicFilter('LastYear');

                // The object returned is a RangeView, not a Range.
                return tableToFilter.getDataBodyRange().getVisibleView().load('values');
            }
        }

        // Queue commands to cache the closed sales data for the selected salesperson
        // as an object type that can later be used to populate a chart: an array 
        // of values.
        function cacheFilteredData(filteredClosedSales) {

            // A RangeView object cannot be assigned as the data source
            // for a chart, so we need to store the values and use them
            // to create a new range, on a hidden worksheet, that will
            // be the data source for the chart on the salesperson's worksheet.
            // First empty closeSales in case the analyze() method has 
            // run before.
            SalesLeadApp.closedSales = [];
            filteredClosedSales.values.forEach(function (item) {
                SalesLeadApp.closedSales.push(item);
            })
        }

        // Queue commands to make a hidden worksheet where the filtered version
        // of the opportunities data will be stored in a range that can be the
        // data source of the chart.
        function createTempSheet(sheets) {
            const tempSheet = sheets.add();
            tempSheet.name = `${SalesLeadApp.selectedSalesPerson}HiddenTemp`;

            // SalesLeadApp.closedSales holds the closed sales from last year
            // for a specific salesperson.
            const range = tempSheet.getRange(`A1:F${SalesLeadApp.closedSales.length}`);
            range.values = SalesLeadApp.closedSales;
            const dateRange = tempSheet.getRange(`E1:E${SalesLeadApp.closedSales.length}`);
            dateRange.numberFormat = 'm/d/yyyy';

            tempSheet.visibility = 'hidden';
        }

        // Queue commands to create the personal worksheet for the selected
        // salesperson, the leads table for the salesperson, and the chart of 
        // his/her closed sales from last year.
        function createAnalysisSheet(sheets, data) {
            const analysisSheet = sheets.add();
            analysisSheet.name = SalesLeadApp.selectedSalesPerson;
            analysisSheet.tabColor = 'green';

            let salesPersonLeadsTable = analysisSheet.tables.add('A3:D3', true);

            // Table names must be unique across the entire workbook, and may not
            // contain spaces.
            let  tableName = `${SalesLeadApp.selectedSalesPerson}LeadsTable`;
            tableName = tableName.replace(" ", "");
            salesPersonLeadsTable.name = tableName;
            salesPersonLeadsTable.getHeaderRowRange().values = [[
                'Account',
                'Probability',
                'Est. Revenue',
                'Expected Value'
            ]];

            const newData = data.map((item, index) => [
                '', // The values of this column will be hyperlinks created below.
                item.Probability,
                item.EstRevenue,
                // Expected value is probability times estimated revenue.
                // The "+ 4" is because arrays are 0-based, but Range addresses 
                // are 1-based and there is a header row, and a gap of 2 rows 
                // at the top of the sheet.
                `=B${index + 4} * C${index + 4}`
            ]);
            salesPersonLeadsTable.rows.add(null, newData);

            // Account names in the table should be hyperlinks to the 
            // customer's row in the Customers worksheet where the user
            // can get customer contact information.
            data.forEach(function (opportunity, index) {
                SalesLeadApp.customerRowAddresses.forEach(function (customerRowAddress) {
                    if (opportunity.Account === customerRowAddress.Name) {
                        const range = analysisSheet.getRange('A' + `${index + 4}`);
                        let hyperlink = {
                            textToDisplay: opportunity.Account,
                            screenTip: 'Click to jump to customer\'s contact information.',
                            documentReference: customerRowAddress.Address
                        }
                        range.hyperlink = hyperlink;
                    }
                });
            })

            salesPersonLeadsTable.columns.getItemAt(2).getRange().numberFormat = [['$#,##0.00']];
            salesPersonLeadsTable.columns.getItemAt(3).getRange().numberFormat = [['$#,##0.00']];

            // Apply a color scale conditonal format to the Expected
            // Value column, so that the lead with the highest potential
            // value is red, the lead with the lowest is blue, and other 
            // leads other leads are mixtures of red, blue, and yellow
            // depending on how close they are to the highest and lowest
            // potential values.
            const rangeToFormat = analysisSheet.getRange(`D3:D${newData.length + 3}`);
            const conditionalFormat = rangeToFormat.conditionalFormats
                .add(Excel.ConditionalFormatType.colorScale);
            const criteria = {
                minimum: {
                    formula: null,
                    type: Excel.ConditionalFormatColorCriterionType.lowestValue,
                    color: '#5858FA' // Light blue
                },
                midpoint: {
                    formula: '1000000',
                    type: Excel.ConditionalFormatColorCriterionType.number,
                    color: '#FFFF00' // Yellow
                },
                maximum: {
                    formula: null,
                    type: Excel.ConditionalFormatColorCriterionType.highestValue,
                    color: '#FA5858' // Light red
                }
            };
            conditionalFormat.colorScale.criteria = criteria;

            // Add a title above the table
            const tableTitleRange = analysisSheet.getRange('B1');
            tableTitleRange.values = [[`${SalesLeadApp.selectedSalesPerson}'s Upcoming Leads`]];
            tableTitleRange.format.font.bold = true;
            tableTitleRange.format.font.size = 16;

            // Make chart of last year's sales for the specific salesperson.
            const hiddenSheet = sheets.getItem(`${SalesLeadApp.selectedSalesPerson}HiddenTemp`);

            // closedSales holds the closed sales from last year for the salesperson.
            const dataRange = hiddenSheet.getRange(`E1:F${SalesLeadApp.closedSales.length}`);
            const chart = analysisSheet.charts.add('ColumnClustered', dataRange, 'auto');

            // Ensure that the chart is to the right of the current leads table.
            chart.setPosition('F1', 'S24')
            chart.title.text = `${SalesLeadApp.selectedSalesPerson}'s Sales Last Year`;
            chart.legend.position = 'right'
            chart.legend.format.fill.setSolidColor('white');
            chart.dataLabels.format.font.size = 15;
            chart.dataLabels.format.font.color = 'black';
            chart.axes.categoryAxis.categoryType = 'DateAxis';
            chart.axes.categoryAxis.title.text = 'Sales Chronologically';
            chart.series.getItemAt(0).name = 'Value in $';
            chart.series.getItemAt(0).trendlines.add("Polynomial");
            let trendline = chart.series.getItemAt(0).trendlines.getItem(0);
            trendline.polynomialOrder = 5;

            // Hide the gridlines on the salesperson's personal sheet.
            analysisSheet.gridlines = false;

            // Auto fit most columns, but not the column with the table title because that
            // makes the column in the table too wide.
            analysisSheet.getRange(`A3:A${newData.length + 3}`).format.autofitColumns();
            analysisSheet.getRange(`C3:D${newData.length + 3}`).format.autofitColumns();
            analysisSheet.getRange(`B3:B${newData.length + 3}`).format.columnWidth = 75;
            analysisSheet.getUsedRange().format.autofitRows();

            analysisSheet.activate();
        }
    }    


    SalesLeadApp.showFileAsByteArray = function (event) {
        const sliceSize = 4096; /*Bytes*/

        // This snippet specifies a small slice size to show how the getFileAsync() method uses slices.
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: sliceSize },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    return onError(result.error);
                }

                // Result.value is the File object.
                SalesLeadApp.getFileContents(result.value, onSuccess, onError);
            });

        function onError(error) {
            SalesLeadApp.errorHandler(error);
        }

        function onSuccess(byteArray) {
            // Now that all of the file content is stored in the "data" parameter,
            // you can do something with it, such as print the file, store the file in a database, etc.
            let base64string = base64js.fromByteArray(byteArray);
            $('#file-contents').val(base64string).show();
        }
    }

    SalesLeadApp.getFileContents = function(file, onSuccess, onError) {
        let expectedSliceCount = file.sliceCount;
        let fileSlices = []; 

        getFileContentsHelper();

        function getFileContentsHelper() {
            file.getSliceAsync(fileSlices.length, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    file.closeAsync();
                    return onError(result.error);
                }

                // Got one slice, store it in a temporary array.
                fileSlices.push(result.value.data);

                if (fileSlices.length == expectedSliceCount) {

                    // All the slices are in the array, so close the file
                    // and concatenate the slices into a single byte array.
                    file.closeAsync();
                    
                    let array = [];
                    fileSlices.forEach(slice => {
                        array = array.concat(slice);
                    });

                    onSuccess(array);
                } else {
                    // We don't have all the slices yet, recursively call 
                    // this method.
                    getFileContentsHelper();
                }
            });
        }
    }

    // Helper function for treating errors
    SalesLeadApp.errorHandler = function (error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        SalesLeadApp.showNotification('Error', error);
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    SalesLeadApp.showNotification = function (content) {
        $('#notification-body').text(content);
        FabricComponents.messageBanner.showBanner();
        FabricComponents.messageBanner.toggleExpansion();
    }

    window.SalesLeadApp = SalesLeadApp;
})();
