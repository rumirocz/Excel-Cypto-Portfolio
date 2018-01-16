# Excel-Cypto-Portfolio

The portfolio I created uses one call to get the entire CoinMarketCap.com ticker and updates your specific coins in your portfolio at times you can set. You can access this from downloading the file "samplePortfolio.xlsm". This file is a macro-enabled Excel spreedsheet. You can download it and alter the data as you see fit. As most won't want to open an untrusted macro-enable file, below are instructions on how to create you own portfolio.

Instructions:

1) Save your Excel workbook as an "Excel-Macro-Enabled Workbook" (.xlsm).
2) Open Visual Basic (On the Developer Ribbon or Alt+F11)
3) Either import the file "coinModule.bas" to your VB Project or Create a new "Module" (renaming is not important) and copy all the text from "coinModule.bas" into the new module.
4) Open the object "ThisWorkbook" in your VBAProject.
5) In the first dropdown box, click on "Workbook".
6) In the second dropdown box, click on "Open".
7) Copy the data from the file "thisWorkbook.bas" into the private sub "Workbook_Open()".
8) Save the modules and VB Project.
9) In your workbook, click on the "Data" ribbon and click "Get Data From Web".
10) In the dialog box, input "https://api.coinmarketcap.com/v1/ticker/?limit=0" and press OK. The Query Editor will open.
11) In the Query Editor, under the ribbon "Transform", click "To Table", don't change any default options, and press OK.
12) The list will have turned into a single column with many rows showing the text "record".
13) Click the icon with two arrows located to the right of the word "Column1".
14) Uncheck the box that says "Use original column name as prefix" and press OK. This will create a query with columns showing everything CoinMarketCap has to offer.
15) Press "Close & Load" on the Home ribbon. This creates a new worksheet with the entire ticker from CoinMarketCap.com.
  15a) On the ribbon, under Data->Queries & Connections, click on "Properties. Then click on the icon to the right of the Query Name text box. This is advanced query properties.
  15b) Check the two boxes that say "Refresh Every" and "Refresh data when opening file". Enter a value for "Refresh Every".
  15c) This will refresh the CoinMarketCap data loaded in your sheet.
16) Rename this sheet to "Ticker".
17) Use the function "cryptoInfo" to link your data to this table.
  17a) For this to work, you click on a sell and enter the formula "=cyrptoInfo(symbol, value)" where symbol = cyrpto symbol and value = CoinMarketData.
  17b) Example: =cryptoInfo("BTC", "price_usd") This will return the current price of Bitcoin in USD.
  17c) Example: =cyrptoInfo("ETH", "price_usd) This will return the current price of Etheruem in USD.
  17d) Example: =cryptoInfo("ADA", "market_cap_usd") This will return the market cap of Cardano in USD.
18) Go back to your VB Project.
19) Open the module file you created or imported.
20) Customize the CONST variables to suit your needs. The names of the sheets have to be the same as the sheets in your workbook or the functions won't return data correctly. You can also set the timer to your liking. This timer is different than the refresh rate of the CoinMarketCap data. The timer here updates/recalculates all the cryptoInfo functions you entered. Save when you are finished. The defaults set here work with my sample workbook file, but may differ from yours.


