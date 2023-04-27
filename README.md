# VBA-challenge
Challenge 2 - Multiple Year Stock Data VBA

The folder contains the following:
 * Easy_Solution Image 
 * Moderate_Solution Image
 * Hard_Solution Image
 * 2019 Hard_Solution Image
 * 2020 Hard_Solution Image
 * Multiple_year_stock_data.bas

VBA formatting was used to loop through stock data from 2018, 2019, and 2020 and output the stock ticker symbol, the yearly change between the opening and closing price, and the percentage annual change, as well as the total stock volume.

The instructor provided a skeleton template for initial setup and two xlsx files with stock data. One file, is a smaller data set that runs the test macros quicker and once completed is used in the larger stock data set for 2018,2019, and 2020.  Images were provided to showcase the desired outcome; however, the file labeled “easy_solution” displays data from 2015 and could not be replicated, due to being outside the scope of given data sets.  Course material, instructor office hours, Study-Group, https://www.excel-easy.com/vba.html, and  https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office were used to create the attached script. 

This macro using loops to run through variables and set criteria to track locations for desired stock data throughout numerous worksheets and appropriately display newly created headers for retrieved data into their appropriately designated cells with the overall goal of identifying the stock with the greatest increase, the greatest decrease, and greatest total volume for 2018, 2019 and 2020.  Conditional formatting highlights positive changes in green and negative changes in red.

In 2018, the stock with the greatest percentage increase is THB with 141.42%, greatest decrease is RKS with -90.02%, and the stock with the greatest total volume is QKN at 1689539560106.

In 2019, the stock with the greatest percentage increase is RYU with 190.03%, greatest decrease is RKS with -91.60%, and the stock with the greatest total volume is ZQD at 4373008528422.

In 2020, the stock with the greatest percentage increase is YDI with 188.76%, greatest decrease is VNG with -89.05%, and the stock with the greatest total volume is QKN at 3452956568861.

Of all the stocks, RKS suffered the greatest percentage decrease in both 2018 and 2019, whereas QKN maintined the greatest total volume in 2018 and 2019.
