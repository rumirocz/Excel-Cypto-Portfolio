# Excel-Cypto-Portfolio

The portfolio I created uses one call to get the entire CoinMarketCap.com ticker and updates your specific coins in your portfolio at times you can set.

This file is a macro-enabled Excel spreedsheet. You can download it and alter the data as you see fit. As most won't want to open an untrusted macro-enable file, below are instructions on how to create you own portfolio.

Instructions:

1) Save your Excel workbook as an "Excel-Macro-Enabled Workbook" (.xlsm).
2) Open Visual Basic (On the Developer Ribbon or Alt+F11)
3) Create a new "Module" (renaming is not important)
4) Copy the data from the file "coinModule.bas" or simply import this module into your Visual Basic project
5) Open the object "ThisWorkbook" in your VBAProject.
6) In the first dropdown box, click on "Workbook".
7) In the second dropdown box, click on "Open".
8) Copy the data from the file "thisWorkbook.bas" into the private sub "Workbook_Open()".
9) Save the modules and you can close Visual Basic
10) In your workbook, click on the "Data" ribbon and click "Get Data From Web".
11) In the dialog box, input "https://api.coinmarketcap.com/v1/ticker/?limit=0" and press OK.
12) In the Query Editor, under the ribbon "Transform", click "ToTable", don't change any default options.
13) The list will have turned into a single column with many rows showing the text "recor".
14) Click the icon with two arrows located to the right of the word "Column1".
15) Uncheck the box that says "Use original column name as prefix" and press OK.
16) Press "Close & Load" on the Home ribbon. This creates a new worksheet with the entire ticker from CoinMarketCap.com
17) Rename this sheet to "Ticker".
18) Use the function "cryptoInfo" to link your data to this table.


