# Excel-Cypto-Portfolio

The portfolio I created uses one call to get the entire CoinMarketCap.com ticker and updates your specific coins in your portfolio at times you can set. You can access this from downloading the file "samplePortfolio.xlsm". This file is a macro-enabled Excel spreedsheet. You can download it and alter the data as you see fit. As most won't want to open an untrusted macro-enable file, below are instructions on how to create you own portfolio.

There is another Excel example on Reddit that I first thought of using, however it forces you to create a query for every single coin you have in your portfolio. If you have dozens or hundreds of coins, this becomes very tedious, and uses alot of overhead when refreshing. You can't just use the full ticker either, the script is required. If you like a cell to the full ticker, when the market caps change, the ranks do as well, and then your cell is linked to a different coin altogether. So use the script. There is also a Google Sheets version of this that someone else made. Much simpler, but Excel has many more options and formatting options than Google Sheets does. The link for the Google Sheets version is https://github.com/rathergood/Crypto-Currency-Price

Instructions:

1) Save your Excel workbook as an "Excel-Macro-Enabled Workbook" (.xlsm).
2) Open Visual Basic (On the Developer Ribbon or Alt+F11)
3) Either import the file "coinModule.bas" to your VB Project or Create a new "Module" (renaming is not important) and copy all the text from "coinModule.bas" into the new module.
4) Open the object "ThisWorkbook" in your VBAProject.
5) Copy the data from the file "thisWorkbook.bas" into the window for ThisWorkbook.
6) This will populate two commands for the workbook.
7) The first Open routine activates the recalculate function and timer and the BeforeClose routine activates the function to disable the timer.
8) Save the modules and VB Project.
9) In your workbook, click on the "Data" ribbon and click "Get Data From Web".
10) In the dialog box, input "https://api.coinmarketcap.com/v1/ticker/?limit=0" and press OK. The Query Editor will open.
  10a) You can change the name of the query by going on the ribbon under "Home"->Query (Properties) if you want. This is helpful if you have multiple queries running.
11) In the Query Editor, under the ribbon "Transform", click "To Table", don't change any default options, and press OK.
12) The list will have turned into a single column with many rows showing the text "record".
13) Click the icon with two arrows located to the right of the word "Column1".
14) Uncheck the box that says "Use original column name as prefix" and press OK. This will create a query with columns showing everything CoinMarketCap has to offer. Don't worry if it says "List Incomplete".
15) Press "Close & Load" on the Home ribbon. This creates a new worksheet with the entire ticker from CoinMarketCap.com. Alternatively, you can click "Close & Load to..." to load the ticker onto an existing worksheet. If doing so, make sure there isn't anything else already in the worksheet as this ticker is 15 columns wide and 1600+ rows deep.
  15a) On the ribbon, under Data->Queries & Connections, click on "Properties. Then click on the icon to the right of the Query Name text box. This is advanced query properties.
  15b) Check the two boxes that say "Refresh Every" and "Refresh data when opening file". Enter a value for "Refresh Every".
  15c) This will refresh the CoinMarketCap data loaded in your sheet when opening the workbook and also every xxx minutes. If you find Excel hanging too often or taking too long to load the query during refresh, consider increasing the amount of time during refreshes. CoinMarketCap could become unresponsive sometimes if they experience high traffic. No values below 1 minute are allowed in Excel.
  15d) Press OK on that dialog and then press OK on the previous dialog to close and save the refresh options.
16) You can rename the ticker sheet to anything you like. This script uses the name "Ticker", so if you don't name your sheet to "Ticker" you will have to edit the script to match (explained at a later step).
17) Use the function "cryptoInfo" to link your data to this table.
  17a) For this to work, you click on a sell and enter the formula "=cyrptoInfo(symbol, value)" where symbol = cyrpto symbol and value = CoinMarketData. These values are not case sensitive but must be an exact match. Crypto symobls are BTC, ETH, LTC, etc. Values are the headings to the columns of the ticker such as price_usd, price_btc, market_cap_usd, etc. Always enclose the text in double quotes.
  17b) Example: =cryptoInfo("BTC", "price_usd") This will return the current price of Bitcoin in USD.
  17c) Example: =cyrptoInfo("ETH", "price_usd) This will return the current price of Etheruem in USD.
  17d) Example: =cryptoInfo("ADA", "market_cap_usd") This will return the market cap of Cardano in USD.
  17e) You could also make a column of just the symbols you own. For instance, column A could contain 20 different coins (just the symbol abbreviations) and then column B could be the cryptoInfo formula dependednt on columnd A. This makes for a quick and easy auto-fill if you have alot of coins. =cryptoInfo(A1, "price_usd"), =crptoInfo(A2, "price_usd), etc.
18) Go back to your VB Project.
19) Open the module file you created or imported.
20) Customize the CONST variables to suit your needs. These are located at the top of the file. The names of the sheets have to be the same as the sheets in your workbook or the functions won't return data correctly. You can also set the timer to your liking. This timer is different than the refresh rate of the CoinMarketCap data. The timer here updates/recalculates all the cryptoInfo functions you entered. Save when you are finished. The defaults set here work with my sample workbook file, but may differ from yours.
  20a) The CONST "sheetToSearch" is the ticker sheet. This is the name of the sheet the ticker was loaded onto. If it's not accurate, the formula won't retrieve any data.
  20b) The CONST "sheetToUpdate" is your worksheet that you entered the cryptoInfo formula on. Change this to be your sheet's name. Currently, I was experiencing some coding issues that the sheet with the formula values wouldn't update when the query refreshes. If you don't change this, your values will always stay the same, even if the ticker updates. If you have multiple sheets that you entered the formula on (more than one sheet that pulls data from the ticker we are using) you will have to add those sheets into the function "timeFunction" at the bottom of the module. Instructions for that are not included in this.
  20c) The CONST "timeToUpdate" is a user defined time for when the previously mentioned sheetToUpdate should take place. The default is every minute, however that may be too often. It should be at the minimum, every time the query refreshes. So if you previously set the query to refresh every 10 minutes, I would set the timeToUpdate to be every 11 minutes. This way it update 1 minute after the query does. The format for this is hh:mm:ss so 02:25:15 would be a refresh once every 2 hours 25 minutues and 15 seconds.
21) Save the VB project and all modules and close them.
22) Any questions let me know, and if something doesn't work, let me know as it took me longer to write this readme than it did the Excel stuff, so the Excel stuff could be buggy.
