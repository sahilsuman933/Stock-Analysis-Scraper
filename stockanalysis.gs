function ImportStockData() {
  var url = "https://stockanalysis.com/api/screener/s/i";
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var json_data = JSON.parse(json);
  var stockData = json_data.data.data;
  var headers = ["Symbol", "Company Name", "Market Cap", "Stock Price", "Price Change 1D (%)", "Industry", "Volume", "PE Ratio"];
  var output = [headers];

  for (var i = 0; i < stockData.length; i++) {
    var item = stockData[i];
    var row = [item.s, item.n, item.marketCap, item.price, item.change, item.industry, item.volume, item.peRatio];
    output.push(row);
  }

  return output;
}

function InitialStockData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var newData = ImportStockData();
  var rangeToUpdate = sheet.getRange(1, 1, newData.length, 8);
  rangeToUpdate.setValues(newData);
  Logger.log("Initial stock data are updated.")

}


var data = data = [{'name': 'Enterprise Value', 'for_value': 'enterpriseValue',
         'link': 'https://stockanalysis.com/api/screener/s/d/enterpriseValue'},
        {'name': 'Market Cap Group', 'for_value': 'marketCapCategory',
         'link': 'https://stockanalysis.com/api/screener/s/d/marketCapCategory'},
        {'name': 'Sector', 'for_value': 'sector', 'link': 'https://stockanalysis.com/api/screener/s/d/sector'},
        {'name': 'Forward PE', 'for_value': 'peForward',
         'link': 'https://stockanalysis.com/api/screener/s/d/peForward'},
        {'name': 'Exchange', 'for_value': 'exchange', 'link': 'https://stockanalysis.com/api/screener/s/d/exchange'},
        {'name': 'Dividend Yield', 'for_value': 'dividendYield',
         'link': 'https://stockanalysis.com/api/screener/s/d/dividendYield'},
        {'name': 'Dividend ($)', 'for_value': 'dps', 'link': 'https://stockanalysis.com/api/screener/s/d/dps'},
        {'name': 'Dividend Growth', 'for_value': 'dividendGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/dividendGrowth'},
        {'name': 'Dividend Growth Years', 'for_value': 'dividendGrowthYears',
         'link': 'https://stockanalysis.com/api/screener/s/d/dividendGrowthYears'},
        {'name': 'Price Change 1W', 'for_value': 'ch1w', 'link': 'https://stockanalysis.com/api/screener/s/d/ch1w'},
        {'name': 'Price Change 1M', 'for_value': 'ch1m', 'link': 'https://stockanalysis.com/api/screener/s/d/ch1m'},
        {'name': 'Price Change 3M', 'for_value': 'ch3m', 'link': 'https://stockanalysis.com/api/screener/s/d/ch3m'},
        {'name': 'Price Change 6M', 'for_value': 'ch6m', 'link': 'https://stockanalysis.com/api/screener/s/d/ch6m'},
        {'name': 'Price Change YTD', 'for_value': 'chYTD', 'link': 'https://stockanalysis.com/api/screener/s/d/chYTD'},
        {'name': 'Price Change 1Y', 'for_value': 'ch1y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch1y'},
        {'name': 'Price Change 3Y', 'for_value': 'ch3y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch3y'},
        {'name': 'Price Change 5Y', 'for_value': 'ch5y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch5y'},
        {'name': 'Price Change 10Y', 'for_value': 'ch10y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch10y'},
        {'name': 'Price Change 15Y', 'for_value': 'ch15y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch15y'},
        {'name': 'Price Change 20Y', 'for_value': 'ch20y', 'link': 'https://stockanalysis.com/api/screener/s/d/ch20y'},
        {'name': 'Price Change 52W Low', 'for_value': 'low52ch',
         'link': 'https://stockanalysis.com/api/screener/s/d/low52ch'},
        {'name': 'Price Change 52W High', 'for_value': 'high52ch',
         'link': 'https://stockanalysis.com/api/screener/s/d/high52ch'},
        {'name': 'All-Time High', 'for_value': 'allTimeHigh',
         'link': 'https://stockanalysis.com/api/screener/s/d/allTimeHigh'},
        {'name': 'All-Time High Change (%)', 'for_value': 'allTimeHighChange',
         'link': 'https://stockanalysis.com/api/screener/s/d/allTimeHighChange'},
        {'name': 'All-Time Low', 'for_value': 'allTimeLow',
         'link': 'https://stockanalysis.com/api/screener/s/d/allTimeLow'},
        {'name': 'All-Time Low Change (%)', 'for_value': 'allTimeLowChange',
         'link': 'https://stockanalysis.com/api/screener/s/d/allTimeLowChange'},
        {'name': 'Analyst Rating', 'for_value': 'analystRatings',
         'link': 'https://stockanalysis.com/api/screener/s/d/analystRatings'},
        {'name': 'Analyst Count', 'for_value': 'analystCount',
         'link': 'https://stockanalysis.com/api/screener/s/d/analystCount'},
        {'name': 'Price Target', 'for_value': 'priceTarget',
         'link': 'https://stockanalysis.com/api/screener/s/d/priceTarget'},
        {'name': 'Price Target Diff. (%)', 'for_value': 'priceTargetChange',
         'link': 'https://stockanalysis.com/api/screener/s/d/priceTargetChange'},
        {'name': 'Open Price', 'for_value': 'open', 'link': 'https://stockanalysis.com/api/screener/s/d/open'},
        {'name': 'Low Price', 'for_value': 'low', 'link': 'https://stockanalysis.com/api/screener/s/d/low'},
        {'name': 'High Price', 'for_value': 'high', 'link': 'https://stockanalysis.com/api/screener/s/d/high'},
        {'name': 'Previous Close', 'for_value': 'close', 'link': 'https://stockanalysis.com/api/screener/s/d/close'},
        {'name': 'Premarket Price', 'for_value': 'premarketPrice',
         'link': 'https://stockanalysis.com/api/screener/s/d/premarketPrice'},
        {'name': 'Premarket % Change', 'for_value': 'premarketChangePercent',
         'link': 'https://stockanalysis.com/api/screener/s/d/premarketChangePercent'},
        {'name': 'After-Hours Price', 'for_value': 'postmarketPrice',
         'link': 'https://stockanalysis.com/api/screener/s/d/postmarketPrice'},
        {'name': 'After-Hours % Change', 'for_value': 'postmarketChangePercent',
         'link': 'https://stockanalysis.com/api/screener/s/d/postmarketChangePercent'},
        {'name': '52 Week Low', 'for_value': 'low52', 'link': 'https://stockanalysis.com/api/screener/s/d/low52'},
        {'name': '52 Week High', 'for_value': 'high52', 'link': 'https://stockanalysis.com/api/screener/s/d/high52'},
        {'name': 'Country', 'for_value': 'country', 'link': 'https://stockanalysis.com/api/screener/s/d/country'},
        {'name': 'Employees', 'for_value': 'employees', 'link': 'https://stockanalysis.com/api/screener/s/d/employees'},
        {'name': 'Empl. Change', 'for_value': 'employeesChange',
         'link': 'https://stockanalysis.com/api/screener/s/d/employeesChange'},
        {'name': 'Empl. Growth', 'for_value': 'employeesChangePercent',
         'link': 'https://stockanalysis.com/api/screener/s/d/employeesChangePercent'},
        {'name': 'Founded', 'for_value': 'founded', 'link': 'https://stockanalysis.com/api/screener/s/d/founded'},
        {'name': 'IPO Date', 'for_value': 'ipoDate', 'link': 'https://stockanalysis.com/api/screener/s/d/ipoDate'},
        {'name': 'Financial Report Date', 'for_value': 'lastReportDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/lastReportDate'},
        {'name': 'Fiscal Year End', 'for_value': 'fiscalYearEnd',
         'link': 'https://stockanalysis.com/api/screener/s/d/fiscalYearEnd'},
        {'name': 'Last 10-K Release Date', 'for_value': 'last10kFilingDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/last10kFilingDate'},
        {'name': 'Revenue', 'for_value': 'revenue', 'link': 'https://stockanalysis.com/api/screener/s/d/revenue'},
        {'name': 'Revenue Growth', 'for_value': 'revenueGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueGrowth'},
        {'name': 'Revenue Growth (Q)', 'for_value': 'revenueGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueGrowthQ'},
        {'name': 'Revenue Growth (3Y)', 'for_value': 'revenueGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueGrowth3Y'},
        {'name': 'Revenue Growth (5Y)', 'for_value': 'revenueGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueGrowth5Y'},
        {'name': 'Gross Profit', 'for_value': 'grossProfit',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossProfit'},
        {'name': 'Gross Profit Growth', 'for_value': 'grossProfitGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossProfitGrowth'},
        {'name': 'Gross Profit Growth (Q)', 'for_value': 'grossProfitGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossProfitGrowthQ'},
        {'name': 'Gross Profit Growth (3Y)', 'for_value': 'grossProfitGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossProfitGrowth3Y'},
        {'name': 'Gross Profit Growth (5Y)', 'for_value': 'grossProfitGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossProfitGrowth5Y'},
        {'name': 'Operating Income', 'for_value': 'operatingIncome',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingIncome'},
        {'name': 'Op. Income Growth', 'for_value': 'operatingIncomeGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingIncomeGrowth'},
        {'name': 'Op. Income Growth (Q)', 'for_value': 'operatingIncomeGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingIncomeGrowthQ'},
        {'name': 'Op. Income Growth (3Y)', 'for_value': 'operatingIncomeGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingIncomeGrowth3Y'},
        {'name': 'Op. Income Growth (5Y)', 'for_value': 'operatingIncomeGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingIncomeGrowth5Y'},
        {'name': 'Net Income', 'for_value': 'netIncome',
         'link': 'https://stockanalysis.com/api/screener/s/d/netIncome'},
        {'name': 'Net Income Growth', 'for_value': 'netIncomeGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/netIncomeGrowth'},
        {'name': 'Net Income Growth (Q)', 'for_value': 'netIncomeGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/netIncomeGrowthQ'},
        {'name': 'Net Income Growth (3Y)', 'for_value': 'netIncomeGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/netIncomeGrowth3Y'},
        {'name': 'Net Income Growth (5Y)', 'for_value': 'netIncomeGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/netIncomeGrowth5Y'},
        {'name': 'Income Tax', 'for_value': 'incomeTax',
         'link': 'https://stockanalysis.com/api/screener/s/d/incomeTax'},
        {'name': 'EPS', 'for_value': 'eps', 'link': 'https://stockanalysis.com/api/screener/s/d/eps'},
        {'name': 'EPS Growth', 'for_value': 'epsGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsGrowth'},
        {'name': 'EPS Growth (Q)', 'for_value': 'epsGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsGrowthQ'},
        {'name': 'EPS Growth (3Y)', 'for_value': 'epsGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsGrowth3Y'},
        {'name': 'EPS Growth (5Y)', 'for_value': 'epsGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsGrowth5Y'},
        {'name': 'EBIT', 'for_value': 'ebit', 'link': 'https://stockanalysis.com/api/screener/s/d/ebit'},
        {'name': 'EBITDA', 'for_value': 'ebitda', 'link': 'https://stockanalysis.com/api/screener/s/d/ebitda'},
        {'name': 'Operating Cash Flow', 'for_value': 'operatingCF',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingCF'},
        {'name': 'Stock-Based Compensation', 'for_value': 'shareBasedComp',
         'link': 'https://stockanalysis.com/api/screener/s/d/shareBasedComp'},
        {'name': 'SBC / Revenue', 'for_value': 'sbcByRevenue',
         'link': 'https://stockanalysis.com/api/screener/s/d/sbcByRevenue'},
        {'name': 'Investing Cash Flow', 'for_value': 'investingCF',
         'link': 'https://stockanalysis.com/api/screener/s/d/investingCF'},
        {'name': 'Financing Cash Flow', 'for_value': 'financingCF',
         'link': 'https://stockanalysis.com/api/screener/s/d/financingCF'},
        {'name': 'Net Cash Flow', 'for_value': 'netCF', 'link': 'https://stockanalysis.com/api/screener/s/d/netCF'},
        {'name': 'Capital expenditures', 'for_value': 'capex',
         'link': 'https://stockanalysis.com/api/screener/s/d/capex'},
        {'name': 'Free Cash Flow', 'for_value': 'fcf', 'link': 'https://stockanalysis.com/api/screener/s/d/fcf'},
        {'name': 'Free Cash Flow - SBC', 'for_value': 'adjustedFCF',
         'link': 'https://stockanalysis.com/api/screener/s/d/adjustedFCF'},
        {'name': 'FCF / Share', 'for_value': 'fcfPerShare',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfPerShare'},
        {'name': 'FCF Growth', 'for_value': 'fcfGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfGrowth'},
        {'name': 'FCF Growth (Q)', 'for_value': 'fcfGrowthQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfGrowthQ'},
        {'name': 'FCF Growth (3Y)', 'for_value': 'fcfGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfGrowth3Y'},
        {'name': 'FCF Growth (5Y)', 'for_value': 'fcfGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfGrowth5Y'},
        {'name': 'Total Cash', 'for_value': 'cash', 'link': 'https://stockanalysis.com/api/screener/s/d/cash'},
        {'name': 'Total Debt', 'for_value': 'debt', 'link': 'https://stockanalysis.com/api/screener/s/d/debt'},
        {'name': 'Debt Growth (YoY)', 'for_value': 'debtGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtGrowth'},
        {'name': 'Debt Growth (QoQ)', 'for_value': 'debtGrowthQoQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtGrowthQoQ'},
        {'name': 'Debt Growth (3Y)', 'for_value': 'debtGrowth3Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtGrowth3Y'},
        {'name': 'Debt Growth (5Y)', 'for_value': 'debtGrowth5Y',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtGrowth5Y'},
        {'name': 'Net Cash / Debt', 'for_value': 'netCash',
         'link': 'https://stockanalysis.com/api/screener/s/d/netCash'},
        {'name': 'Net Cash Growth', 'for_value': 'netCashGrowth',
         'link': 'https://stockanalysis.com/api/screener/s/d/netCashGrowth'},
        {'name': 'Net Cash / Market Cap', 'for_value': 'netCashByMarketCap',
         'link': 'https://stockanalysis.com/api/screener/s/d/netCashByMarketCap'},
        {'name': 'Assets', 'for_value': 'assets', 'link': 'https://stockanalysis.com/api/screener/s/d/assets'},
        {'name': 'Liabilities', 'for_value': 'liabilities',
         'link': 'https://stockanalysis.com/api/screener/s/d/liabilities'},
        {'name': 'Gross Margin', 'for_value': 'grossMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/grossMargin'},
        {'name': 'Operating Margin', 'for_value': 'operatingMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/operatingMargin'},
        {'name': 'Profit Margin', 'for_value': 'profitMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/profitMargin'},
        {'name': 'FCF Margin', 'for_value': 'fcfMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/fcfMargin'},
        {'name': 'EBITDA Margin', 'for_value': 'ebitdaMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/ebitdaMargin'},
        {'name': 'EBIT Margin', 'for_value': 'ebitMargin',
         'link': 'https://stockanalysis.com/api/screener/s/d/ebitMargin'},
        {'name': 'Research & Development', 'for_value': 'researchAndDevelopment',
         'link': 'https://stockanalysis.com/api/screener/s/d/researchAndDevelopment'},
        {'name': 'R&D / Revenue', 'for_value': 'rndByRevenue',
         'link': 'https://stockanalysis.com/api/screener/s/d/rndByRevenue'},
        {'name': 'PS Ratio', 'for_value': 'psRatio', 'link': 'https://stockanalysis.com/api/screener/s/d/psRatio'},
        {'name': 'Forward PS', 'for_value': 'psForward',
         'link': 'https://stockanalysis.com/api/screener/s/d/psForward'},
        {'name': 'PB Ratio', 'for_value': 'pbRatio', 'link': 'https://stockanalysis.com/api/screener/s/d/pbRatio'},
        {'name': 'P/FCF Ratio', 'for_value': 'pFcfRatio',
         'link': 'https://stockanalysis.com/api/screener/s/d/pFcfRatio'},
        {'name': 'PEG Ratio', 'for_value': 'pegRatio', 'link': 'https://stockanalysis.com/api/screener/s/d/pegRatio'},
        {'name': 'EV/Sales', 'for_value': 'evSales', 'link': 'https://stockanalysis.com/api/screener/s/d/evSales'},
        {'name': 'Forward EV/Sales', 'for_value': 'evSalesForward',
         'link': 'https://stockanalysis.com/api/screener/s/d/evSalesForward'},
        {'name': 'EV/Earnings', 'for_value': 'evEarnings',
         'link': 'https://stockanalysis.com/api/screener/s/d/evEarnings'},
        {'name': 'EV/EBITDA', 'for_value': 'evEbitda', 'link': 'https://stockanalysis.com/api/screener/s/d/evEbitda'},
        {'name': 'EV/EBIT', 'for_value': 'evEbit', 'link': 'https://stockanalysis.com/api/screener/s/d/evEbit'},
        {'name': 'EV/FCF', 'for_value': 'evFcf', 'link': 'https://stockanalysis.com/api/screener/s/d/evFcf'},
        {'name': 'Earnings Yield', 'for_value': 'earningsYield',
         'link': 'https://stockanalysis.com/api/screener/s/d/earningsYield'},
        {'name': 'FCF Yield', 'for_value': 'fcfYield', 'link': 'https://stockanalysis.com/api/screener/s/d/fcfYield'},
        {'name': 'Payout Ratio', 'for_value': 'payoutRatio',
         'link': 'https://stockanalysis.com/api/screener/s/d/payoutRatio'},
        {'name': 'Payout Frequency', 'for_value': 'payoutFrequency',
         'link': 'https://stockanalysis.com/api/screener/s/d/payoutFrequency'},
        {'name': 'Buyback Yield', 'for_value': 'buybackYield',
         'link': 'https://stockanalysis.com/api/screener/s/d/buybackYield'},
        {'name': 'Shareholder Yield', 'for_value': 'totalReturn',
         'link': 'https://stockanalysis.com/api/screener/s/d/totalReturn'},
        {'name': 'Average Volume', 'for_value': 'averageVolume',
         'link': 'https://stockanalysis.com/api/screener/s/d/averageVolume'},
        {'name': 'Relative Volume', 'for_value': 'relativeVolume',
         'link': 'https://stockanalysis.com/api/screener/s/d/relativeVolume'},
        {'name': 'Beta (1Y)', 'for_value': 'beta', 'link': 'https://stockanalysis.com/api/screener/s/d/beta'},
        {'name': 'Relative Strength Index (RSI)', 'for_value': 'rsi',
         'link': 'https://stockanalysis.com/api/screener/s/d/rsi'}, {'name': 'Short % Float', 'for_value': 'shortFloat',
                                                                     'link': 'https://stockanalysis.com/api/screener/s/d/shortFloat'},
        {'name': 'Short % Shares', 'for_value': 'shortShares',
         'link': 'https://stockanalysis.com/api/screener/s/d/shortShares'},
        {'name': 'Short Ratio', 'for_value': 'shortRatio',
         'link': 'https://stockanalysis.com/api/screener/s/d/shortRatio'},
        {'name': 'Shares Out', 'for_value': 'sharesOut',
         'link': 'https://stockanalysis.com/api/screener/s/d/sharesOut'},
        {'name': 'Float', 'for_value': 'float', 'link': 'https://stockanalysis.com/api/screener/s/d/float'},
        {'name': 'Shares Ch. (YoY)', 'for_value': 'sharesYoY',
         'link': 'https://stockanalysis.com/api/screener/s/d/sharesYoY'},
        {'name': 'Shares Ch. (QoQ)', 'for_value': 'sharesQoQ',
         'link': 'https://stockanalysis.com/api/screener/s/d/sharesQoQ'},
        {'name': 'Shares Insiders', 'for_value': 'sharesInsiders',
         'link': 'https://stockanalysis.com/api/screener/s/d/sharesInsiders'},
        {'name': 'Shares Institut.', 'for_value': 'sharesInstitutions',
         'link': 'https://stockanalysis.com/api/screener/s/d/sharesInstitutions'},
        {'name': 'Earnings Date', 'for_value': 'earningsDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/earningsDate'},
        {'name': 'Is SPAC', 'for_value': 'isSpac', 'link': 'https://stockanalysis.com/api/screener/s/d/isSpac'},
        {'name': 'Ex-Div Date', 'for_value': 'exDivDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/exDivDate'},
        {'name': 'Payment Date', 'for_value': 'paymentDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/paymentDate'},
        {'name': 'Return on Equity', 'for_value': 'roe', 'link': 'https://stockanalysis.com/api/screener/s/d/roe'},
        {'name': 'Return on Assets', 'for_value': 'roa', 'link': 'https://stockanalysis.com/api/screener/s/d/roa'},
        {'name': 'Return on Capital', 'for_value': 'roic', 'link': 'https://stockanalysis.com/api/screener/s/d/roic'},
        {'name': 'Return on Equity (5Y)', 'for_value': 'roe5y',
         'link': 'https://stockanalysis.com/api/screener/s/d/roe5y'},
        {'name': 'Return on Assets (5Y)', 'for_value': 'roa5y',
         'link': 'https://stockanalysis.com/api/screener/s/d/roa5y'},
        {'name': 'Return on Capital (5Y)', 'for_value': 'roic5y',
         'link': 'https://stockanalysis.com/api/screener/s/d/roic5y'},
        {'name': 'Rev / Employee', 'for_value': 'revPerEmployee',
         'link': 'https://stockanalysis.com/api/screener/s/d/revPerEmployee'},
        {'name': 'Prof. / Employee', 'for_value': 'profitPerEmployee',
         'link': 'https://stockanalysis.com/api/screener/s/d/profitPerEmployee'},
        {'name': 'Asset Turnover', 'for_value': 'assetTurnover',
         'link': 'https://stockanalysis.com/api/screener/s/d/assetTurnover'},
        {'name': 'Inventory Turnover', 'for_value': 'inventoryTurnover',
         'link': 'https://stockanalysis.com/api/screener/s/d/inventoryTurnover'},
        {'name': 'Current Ratio', 'for_value': 'currentRatio',
         'link': 'https://stockanalysis.com/api/screener/s/d/currentRatio'},
        {'name': 'Quick Ratio', 'for_value': 'quickRatio',
         'link': 'https://stockanalysis.com/api/screener/s/d/quickRatio'},
        {'name': 'Debt / Equity', 'for_value': 'debtEquity',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtEquity'},
        {'name': 'Debt / EBITDA', 'for_value': 'debtEbitda',
         'link': 'https://stockanalysis.com/api/screener/s/d/debtEbitda'},
        {'name': 'Debt / FCF', 'for_value': 'debtFcf', 'link': 'https://stockanalysis.com/api/screener/s/d/debtFcf'},
        {'name': 'Effective Tax Rate', 'for_value': 'taxRate',
         'link': 'https://stockanalysis.com/api/screener/s/d/taxRate'},
        {'name': 'Tax / Revenue', 'for_value': 'taxByRevenue',
         'link': 'https://stockanalysis.com/api/screener/s/d/taxByRevenue'},
        {'name': "Shareholders' Equity", 'for_value': 'equity',
         'link': 'https://stockanalysis.com/api/screener/s/d/equity'},
        {'name': 'Working Capital', 'for_value': 'workingCapital',
         'link': 'https://stockanalysis.com/api/screener/s/d/workingCapital'},
        {'name': 'Last Stock Split', 'for_value': 'lastSplitType',
         'link': 'https://stockanalysis.com/api/screener/s/d/lastSplitType'},
        {'name': 'Last Split Date', 'for_value': 'lastSplitDate',
         'link': 'https://stockanalysis.com/api/screener/s/d/lastSplitDate'},
        {'name': 'Altman Z-Score', 'for_value': 'zScore', 'link': 'https://stockanalysis.com/api/screener/s/d/zScore'},
        {'name': 'Piotroski F-Score', 'for_value': 'fScore',
         'link': 'https://stockanalysis.com/api/screener/s/d/fScore'},
        {'name': 'EPS Growth This Quarter', 'for_value': 'epsThisQuarter',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsThisQuarter'},
        {'name': 'EPS Growth Next Quarter', 'for_value': 'epsNextQuarter',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsNextQuarter'},
        {'name': 'EPS Growth This Year', 'for_value': 'epsThisYear',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsThisYear'},
        {'name': 'EPS Growth Next Year', 'for_value': 'epsNextYear',
         'link': 'https://stockanalysis.com/api/screener/s/d/epsNextYear'},
        {'name': 'Revenue Growth This Quarter', 'for_value': 'revenueThisQuarter',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueThisQuarter'},
        {'name': 'Revenue Growth Next Quarter', 'for_value': 'revenueNextQuarter',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueNextQuarter'},
        {'name': 'Revenue Growth This Year', 'for_value': 'revenueThisYear',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueThisYear'},
        {'name': 'Revenue Growth Next Year', 'for_value': 'revenueNextYear',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenueNextYear'},
        {'name': 'EPS Growth Next 5Y', 'for_value': 'eps5y',
         'link': 'https://stockanalysis.com/api/screener/s/d/eps5y'},
        {'name': 'Revenue Growth Next 5Y', 'for_value': 'revenue5y',
         'link': 'https://stockanalysis.com/api/screener/s/d/revenue5y'},
        {'name': 'In Index', 'for_value': 'inIndex', 'link': 'https://stockanalysis.com/api/screener/s/d/inIndex'},
        {'name': 'Views', 'for_value': 'views', 'link': 'https://stockanalysis.com/api/screener/s/d/views'},
        {'name': 'Interest Coverage Ratio', 'for_value': 'interestCoverage',
         'link': 'https://stockanalysis.com/api/screener/s/d/interestCoverage'},
        {'name': 'Return From IPO Price', 'for_value': 'ipr', 'link': 'https://stockanalysis.com/api/screener/s/d/ipr'},
        {'name': 'Return From Open Price', 'for_value': 'iprfo',
         'link': 'https://stockanalysis.com/api/screener/s/d/iprfo'},
        {'name': '50-Day Moving Average', 'for_value': 'ma50',
         'link': 'https://stockanalysis.com/api/screener/s/d/ma50'},
        {'name': '200-Day Moving Average', 'for_value': 'ma200',
         'link': 'https://stockanalysis.com/api/screener/s/d/ma200'},
        {'name': 'Price Change 50-Day MA', 'for_value': 'ma50ch',
         'link': 'https://stockanalysis.com/api/screener/s/d/ma50ch'},
        {'name': 'Price Change 200-Day MA', 'for_value': 'ma200ch',
         'link': 'https://stockanalysis.com/api/screener/s/d/ma200ch'}]








function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Stock Analysis');
  menu.addItem('Insert Stock Analysis Data', 'InsertStockData');
  menu.addToUi();
}

function InsertAllFilterData() {
  var requests = data.map(function(item) {
    return {url: item.link, muteHttpExceptions: true};
  });

  var responses = UrlFetchApp.fetchAll(requests);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var symbolsRange = sheet.getDataRange(); 
  var symbolsData = symbolsRange.getValues();
  var startColumn = 9;
  var allData = [];

  responses.forEach(function(response, index) {
    var jsonResponse = JSON.parse(response.getContentText());
    var endpointData = jsonResponse.data.data;
    var output = processEndpointData(endpointData, symbolsData, data[index].name);

    output.forEach(function(row, rowIndex) {
      if (!allData[rowIndex]) {
        allData[rowIndex] = [];
      }
      allData[rowIndex].push(...row);
    });
  });

  var outputRange = sheet.getRange(1, startColumn, allData.length, allData[0].length);
  outputRange.setValues(allData);

  Logger.log("Filter data are updated.")
}

function processEndpointData(endpointData, symbolsData, headerName) {
  var output = [[headerName]];

  for (var i = 1; i < symbolsData.length; i++) {
    var symbol = symbolsData[i][0];
    var matchedData = endpointData.find(row => row[0] === symbol);
    var valueToWrite = matchedData ? matchedData[1] : "-";
    output.push([valueToWrite]);
  }

  return output;
}


function InsertStockData(){
  InitialStockData()
  InsertAllFilterData()
}
