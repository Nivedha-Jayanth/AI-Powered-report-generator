// --- Configuration ---
const RAW_DATA_SHEET_NAME = "Sales Data"; // Name of the sheet containing raw sales data
const REPORT_OUTPUT_SHEET_NAME = "Monthly Report"; // Name of the sheet to output the report and charts
const GEMINI_MODEL = "gemini-1.5-flash"; // You can also try "gemini-1.5-pro" for potentially more detailed responses

// Function to get the API key securely from script properties
function getGeminiApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("Gemini API Key not found in Script Properties. Please add it as 'GEMINI_API_KEY'.");
  }
  return apiKey;
}

// Function to make API call to Gemini
function callGemini(prompt) {
  const apiKey = getGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.7, // Adjust for creativity (0.0 - 1.0). Higher values = more creative, lower = more focused.
      maxOutputTokens: 800, // Adjust based on expected report length.
    },
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Allows inspection of error responses
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data.candidates && data.candidates.length > 0 && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts.length > 0) {
      return data.candidates[0].content.parts[0].text;
    } else if (data.error) {
      Logger.log("Gemini API Error: " + JSON.stringify(data.error));
      return "Error: " + data.error.message;
    } else {
      Logger.log("Unexpected Gemini API response: " + JSON.stringify(data));
      return "Error: Could not generate report. Unexpected API response.";
    }
  } catch (e) {
    Logger.log("Error calling Gemini API: " + e.toString());
    return "Error: Failed to connect to Gemini API. Check logs for details.";
  }
}

// Function to get sales data from the spreadsheet
function getSalesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error", `Sheet named '${RAW_DATA_SHEET_NAME}' not found. Please create it or update the script.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) { // Checks if there's only header or no data
    SpreadsheetApp.getUi().alert("Info", `No sales data found in '${RAW_DATA_SHEET_NAME}' beyond the header row.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }

  // Assuming first row is header
  const headers = values[0].map(header => String(header).trim()); // Trim headers to avoid whitespace issues
  const salesData = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[i][j];
    }
    salesData.push(row);
  }
  return salesData;
}

// Function to generate and write the monthly report
function generateMonthlySalesReport() {
  const ui = SpreadsheetApp.getUi();

  const salesData = getSalesData();
  if (!salesData || salesData.length === 0) {
    return; // Exit if no data or error
  }

  // --- Data Aggregation for AI Report and Charts ---
  const productSales = {};
  const regionSales = {};
  let totalSales = 0;

  salesData.forEach(row => {
    const product = row['Product'];
    const region = row['Region'];
    const salesAmount = parseFloat(row['Sales Amount (INR)']); // Ensure sales amount is a number

    if (!isNaN(salesAmount)) {
      productSales[product] = (productSales[product] || 0) + salesAmount;
      regionSales[region] = (regionSales[region] || 0) + salesAmount;
      totalSales += salesAmount;
    }
  });

  // Sort products and regions by sales for top/low performers
  const sortedProducts = Object.entries(productSales).sort(([, a], [, b]) => b - a);
  const sortedRegions = Object.entries(regionSales).sort(([, a], [, b]) => b - a);

  // Prepare data for LLM prompt
  let dataSummaryForLLM = "Here is a summary of sales performance:\n\n";
  dataSummaryForLLM += "Product Performance (Product: Total Sales):\n";
  sortedProducts.forEach(([product, sales]) => {
    dataSummaryForLLM += `- ${product}: ${sales.toFixed(2)} INR\n`;
  });
  dataSummaryForLLM += "\nRegion Performance (Region: Total Sales):\n";
  sortedRegions.forEach(([region, sales]) => {
    dataSummaryForLLM += `- ${region}: ${sales.toFixed(2)} INR\n`;
  });
  dataSummaryForLLM += `\nOverall Total Sales: ${totalSales.toFixed(2)} INR\n\n`;

  // Add a sample of raw data for context, if the dataset is not overwhelmingly large
  const numRowsToShowAsSample = Math.min(salesData.length, 20); // Limit raw data sample to 20 rows
  if (numRowsToShowAsSample > 0) {
      dataSummaryForLLM += "Sample of raw data for additional context:\n";
      dataSummaryForLLM += "Product,Region,Date,Sales Amount (INR)\n";
      dataSummaryForLLM += salesData.slice(0, numRowsToShowAsSample).map(row => {
        // Handle potential missing headers or varied casing
        const product = row['Product'] || row.product || '';
        const region = row['Region'] || row.region || '';
        const date = row['Date'] instanceof Date ? row['Date'].toISOString().split('T')[0] : row['Date'] || ''; // Format date nicely
        const sales = row['Sales Amount (INR)'] || row.Sales || row.sales || '';
        return `${product},${region},${date},${sales}`;
      }).join("\n");
      if (salesData.length > numRowsToShowAsSample) {
          dataSummaryForLLM += `\n...and ${salesData.length - numRowsToShowAsSample} more raw sales records.`;
      }
  }

  const prompt = `Based on the following sales data:\n\n${dataSummaryForLLM}\n\nGenerate a concise monthly sales report. The report should include:
- A brief overview of the sales period.
- Top 3-5 performing products based on sales amount.
- An analysis of region performance, highlighting regions with high and low sales.
- Any notable insights, trends, or observations from the data (e.g., specific product doing well in a region, overall growth/decline).
Format the report clearly with headings and bullet points where appropriate. Aim for a professional and easy-to-read summary.`;

  ui.alert("Generating Report...", "Please wait while Gemini analyzes the sales data and generates the report. This may take a moment.", ui.ButtonSet.OK);

  const reportText = callGemini(prompt);

  if (reportText.startsWith("Error:")) {
    ui.alert("Report Generation Failed", reportText, ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportSheet = ss.getSheetByName(REPORT_OUTPUT_SHEET_NAME);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(REPORT_OUTPUT_SHEET_NAME);
  } else {
    reportSheet.clearContents(); // Clear previous report and charts
    // Remove existing charts if any, to avoid duplicates
    const charts = reportSheet.getCharts();
    charts.forEach(chart => reportSheet.removeChart(chart));
  }

  // --- Write AI Report to Sheet ---
  let currentRow = 1;
  reportSheet.getRange(currentRow++, 1).setValue("Monthly Sales Report - Generated by AI");
  reportSheet.getRange(currentRow++, 1).setValue("Generated On: " + new Date().toLocaleString());
  currentRow++; // Blank line
  reportSheet.getRange(currentRow++, 1).setValue(reportText);

  // Auto-resize column for report text
  reportSheet.autoResizeColumn(1);
  reportSheet.setColumnWidth(1, 800); // Set a reasonable width for the report column

  currentRow += 2; // Add some space after the text report

  // --- Prepare Data for Charts in a hidden area or near charts ---
  // For charting, it's best to put the data directly on the sheet.
  // We'll put it starting from column 3 to keep it separated from the main report text.
  const chartDataStartRow = currentRow;
  const chartDataStartCol = 3; // Start data for charts in column C

  // Product Sales Data for Chart
  reportSheet.getRange(chartDataStartRow, chartDataStartCol).setValue("Product");
  reportSheet.getRange(chartDataStartRow, chartDataStartCol + 1).setValue("Total Sales (INR)");
  sortedProducts.forEach((item, index) => {
    reportSheet.getRange(chartDataStartRow + 1 + index, chartDataStartCol).setValue(item[0]);
    reportSheet.getRange(chartDataStartRow + 1 + index, chartDataStartCol + 1).setValue(item[1]);
  });
  const productDataRange = reportSheet.getRange(chartDataStartRow, chartDataStartCol, sortedProducts.length + 1, 2);

  // Region Sales Data for Chart (start below product data + some space)
  const regionDataStartRow = chartDataStartRow + sortedProducts.length + 3; // Space after product data
  reportSheet.getRange(regionDataStartRow, chartDataStartCol).setValue("Region");
  reportSheet.getRange(regionDataStartRow, chartDataStartCol + 1).setValue("Total Sales (INR)");
  sortedRegions.forEach((item, index) => {
    reportSheet.getRange(regionDataStartRow + 1 + index, chartDataStartCol).setValue(item[0]);
    reportSheet.getRange(regionDataStartRow + 1 + index, chartDataStartCol + 1).setValue(item[1]);
  });
  const regionDataRange = reportSheet.getRange(regionDataStartRow, chartDataStartCol, sortedRegions.length + 1, 2);

  // --- Generate Charts ---

  // Product Sales Chart
  const productChart = reportSheet.newChart()
      .asColumnChart() // Or .asBarChart() if you prefer horizontal bars
      .addRange(productDataRange)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Total Sales by Product')
      .setOption('hAxis.title', 'Product')
      .setOption('vAxis.title', 'Sales Amount (INR)')
      .setOption('legend.position', 'none') // No need for legend if only one series
      .setOption('width', 600)
      .setOption('height', 400)
      .setPosition(currentRow, 1, 0, 0) // Position below the text report
      .build();
  reportSheet.insertChart(productChart);

  
  // Region Sales Chart
  const regionChart = reportSheet.newChart()
      .asColumnChart() // Or .asBarChart()
      .addRange(regionDataRange)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setOption('title', 'Total Sales by Region')
      .setOption('hAxis.title', 'Region')
      .setOption('vAxis.title', 'Sales Amount (INR)')
      .setOption('legend.position', 'none')
      .setOption('width', 600)
      .setOption('height', 400)
      .setPosition(currentRow, 1, 0, 0) // Position below the previous chart
      .build();
  reportSheet.insertChart(regionChart);

  // Auto-resize columns for chart data (optional, they are often hidden or in a separate tab)
  reportSheet.autoResizeColumn(chartDataStartCol);
  reportSheet.autoResizeColumn(chartDataStartCol + 1);

  ui.alert("Report & Charts Generated!", `The monthly sales report and charts have been successfully generated in the '${REPORT_OUTPUT_SHEET_NAME}' sheet.`, ui.ButtonSet.OK);
}

// Function to add a custom menu item to the spreadsheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Reports')
      .addItem('Generate Monthly Sales Report & Charts', 'generateMonthlySalesReport')
      .addToUi();
}
