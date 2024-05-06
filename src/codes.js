/**
 * Integrates Notion and Google Spreadsheets and creates a pie chart based on the category for each database.
 *
*/
class ChartManager {
     /**
     * @constructor
     * @param {string} sheetNameDatabase - The name of the sheet where the database is located.
     * @param {string} tableId - The ID of the table.
     * @param {string} itemName - The name of the item.
     * @param {string} categoryName - The name of the category.
     * @param {string} payName - The name of the pay.
     * @param {string} chartName - The name of the chart.
     * @param {number} row - The row where the chart is located.
     * @param {number} column - The column where the chart is located.
     */
    constructor(sheetNameDatabase,tableId,itemName,categoryName,payName,chartName,row,column) {
        this.sheetDataBase = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameDatabase);
        this.tableId = tableId;
        this.itemName = itemName;
        this.categoryName = categoryName;
        this.payName = payName;
        this.chartName = chartName;
        this.row = row;
        this.column = column;
    }
      
      /**
       * Private method to input information from the designated database in Notion
       * Used in createPieChart method
       * @return {object} sumDataRange - The range of the data to be used for the chart.
      */
      inputDatabase_() {
        // console.log(this.tableId,this.itemName,this.categoryName);
        let url = `https://api.notion.com/v1/databases/${this.tableId}/query`;
        let options = {
          method: 'post',
          headers: {
            'Authorization': 'Bearer ' + notionToken,  
            'Notion-Version': '2021-08-16'  
          },
        };
        
        // Get the database information from Notion.
        let response = UrlFetchApp.fetch(url, options);
        let databaseInfo = JSON.parse(response.getContentText());
        let results = databaseInfo.results;
        // console.log(results);

        // Get the designated data from the database and store as array format.
        let allData = [];

        for (let i = 0; i < results.length; i++) {
          let item;
          let category;
          let pay;
          // Determine item based on type
          switch (results[i].properties[this.itemName].type) {
            case 'title':
                // console.log("Data type is title");
                item = results[i].properties[this.itemName]?.title?.[0]?.text?.content;
                break;
            default:
                // console.log("Data type is undefined, probably data is not input");
                item = undefined;
          }

          // Determine category based on type
          switch (results[i].properties[this.categoryName].type) {
            case 'select':
                // console.log("Data type is select");
                category = results[i].properties[this.categoryName]?.select?.name;
                break;
            case 'relation':
                // console.log("Data type is select");
                category = this.getRelationValue_(results[i],this.categoryName);
                break;
            default:
                // console.log("Data type is undefined, probably data is not input");
                category = undefined;
          }

          // Determine number based on type
          switch (results[i].properties[this.payName].type) {
            case 'number':
                // console.log("Data type is number");
                pay = results[i].properties[this.payName]?.number;
                break;            
            default:
                // console.log("Data type is undefined, probably data is not input");
                pay = undefined;
          }

          // Checks whether the target record is the one to be input into the spreadsheet or not. If not, skip the record.
          if (item === undefined){
            console.log(`Item data for record ${i} is undefined: ${results[i].properties[this.itemName]?.title?.[0]?.text?.content}`);
            continue;
          } else if (category === undefined){ 
            console.log(`Category data for record ${i} is undefined: ${results[i].properties[this.categoryName]?.select?.name}`);
            continue;
          } else if (pay === undefined){
            console.log(`Pay data for ${item} is undefined: ${results[i].properties[this.payName].number}`);
            continue;
          } else if (pay === null){
            console.log(`Pay data for record ${item} is null: ${results[i].properties[this.payName].number}`);
          } else if (pay === undefined){
            console.log(`Pay data for record ${item} is undefined: ${results[i].properties[this.payName].number}`);
          } else if (pay === 0){
            console.log(`Pay data for record ${item} is 0`);
            continue;
          }

          allData.push([item,category,pay]);
        }
        console.log(allData);
        
        //Checks if the data is already in the spreadsheet. If so, delete the data.
        const lastRow = this.sheetDataBase.getLastRow();
        if (lastRow > 1){
            this.sheetDataBase.deleteRows(2,lastRow-1);
        }
        
        // Input the data to the spreadsheet.
        const allDataRange = this.sheetDataBase.getRange(2,1,allData.length,3);
        allDataRange.setValues(allData);

        const refData = this.sheetDataBase.getRange(2, 2, lastRow - 1, 2).getValues();

        // Sum the monthly pay for each category
        const totals = {};
        refData.forEach(([category, pay]) => {
          let categoryArray = category.split(", ");
          for (let i = 0; i < categoryArray.length; i++){
            totals[categoryArray[i]] = (totals[categoryArray[i]] || 0) + Number(pay);
          }
        });
        
        // Convert the totals to a 2D array suitable for a chart
        const chartData = [['Category', 'Total Pay']];
        for (let category in totals) {
          chartData.push([category, totals[category]]);
        }
        
        // Add the chart data to the sheet
        const sumDataRange = this.sheetDataBase.getRange(1, 5, chartData.length, 2);
        sumDataRange.setValues(chartData);

        return sumDataRange;
        }

      /**
       * Private method to to retrieve value of the relation of the database in Notion
       * Used in inputDatabase_ method
       * @param {object} result - The result of the database.
       * @param {string} relationName - The name of the relation.
       * @return {Array} pageContents - The value of the relation.
      */
      getRelationValue_(result,relationName) {
          const relationIds = result.properties[relationName].relation; // Makes the variable name plural in a case that there are multiple categories for one item
          switch (relationIds){
            case 'undefined':
              relationData = "undefined"
              break;
            case '[]':
              relationData = "empty array"
              break;
            default:
              let pageContents = [];
              for (let relationId of relationIds) {
                let pageUrl = `https://api.notion.com/v1/pages/${relationId.id}`;
                let pageOptions = {
                  method: 'get',
                  headers: {
                    'Authorization': 'Bearer ' + notionToken,  
                    'Notion-Version': '2021-08-16'  
                  },
                };
                let pageResponse = UrlFetchApp.fetch(pageUrl, pageOptions);
                let pageInfo = JSON.parse(pageResponse.getContentText());
                let pageContent = pageInfo.properties[this.categoryName].title[0].plain_text.trim();
                pageContents.push(pageContent);
              }
              if (relationIds.length >1){
                console.log(`Multiple categories: ${pageContents}`);
                pageContents = [pageContents.join(", ")];
              }
              return pageContents;
          }
      }

      //method to create pie chart from that database in Notion on the sheet
      createPieChart() {
        const sumDataRange = this.inputDatabase_();

        //Checks chart creation status and do a different process depending on whether there is a chart or not
        const sheetChart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameChart);
        if (sheetChart.getCharts() === null){
            console.log("There is no chart in the sheet. Will create a new chart.");
            // Builds and insert the chart
            chartBuilder = sheetChart.newChart()
                .setChartType(Charts.ChartType.PIE)
                .addRange(sumDataRange)
                .setOption('title',`${this.chartName}`)
                .setPosition(this.row, this.column, 0, 0)
                .build();
            sheetChart.insertChart(chartBuilder);
        } else {
            const charts = sheetChart.getCharts();
            const chartNames = charts.map(chart => chart.getOptions().get('title'));
            //I want to divide execution depending on whether the chart with a certain name exists
            const checkTableExist = chartNames.indexOf(this.chartName);
            let chartBuilder;
            if (checkTableExist === -1){
                console.log("There is no chart with the target chart name in the sheet. Will create a new chart.");
                // Builds and insert the chart
                chartBuilder = sheetChart.newChart()
                    .setChartType(Charts.ChartType.PIE)
                    .addRange(sumDataRange)
                    .setOption('title',`${this.chartName}`)
                    .setOption('titleTextStyle', {fontSize: 24})
                    .setPosition(this.row, this.column, 0, 0)
                    .build();
                sheetChart.insertChart(chartBuilder);
            } else {
                console.log("There is a previous chart in the sheet. Will update the existing chart.");
                //Replaces the previous data with the new one in the existing chart
                let chartIndex = chartNames.indexOf(this.chartName);
                let chart = charts[chartIndex];
                chartBuilder = chart.modify()
                .setChartType(Charts.ChartType.PIE)
                .addRange(sumDataRange)
                .build();
                sheetChart.updateChart(chartBuilder);
            }
        }
    }
}

/*
  Instances to manage the pie chart for target databases with trigger when the spreadsheet is opened
*/
function manageChart_1(){
  const chart_1 = new ChartManager (sheetName_1,tableId_1,item_1,category_1,monthlyPay_1,chartName_1,2,1);
  chart_1.createPieChart();
}

function manageChart_2(){
  const chart_2 = new ChartManager (sheetName_2,tableId_2,item_2,category_2,pay_2,chartName_2,2,7);
  chart_2.createPieChart();
}

function manageChart_3(){
  const chart_3 = new ChartManager (sheetName_3,tableId_3,item_3,category_3,pay_3,chartName_3,1,18);
  chart_3.createPieChart();
}