# vba-challenge

Summary of the stocks for three consecutive years from 2014 to 2016. 

*Data was acquired from the UCSD-Extension bootcamp on Data Analysis and Visualization, and code was written for a graded homework*
___________
## Excel file structure
The data should be in the following structure, and the left-top corner must be at cell "A2" of the worksheet, just below the header
| \<ticker\> | \<date\> | \<open\> | \<high\> | \<low\> | \<close\> | \<vol\> |
|----------|--------|-------|-------|-------|---------|--------|

\<ticker\>: the unique identifier for the stock
\<date\>: day of the records
\<open\>: opening stock price
\<high\>: highest stock price of the day
\<low\>: lowest stock price of the day
\<close\>: volume of the stock of the day

### Alphabetical:
The file alphabeticStatistics.bas prefers a workbook divided into many worksheets based on the first letter of the \<ticker\>. The macro combines all the worksheets into one, and prints the summary to the first worksheet.
### Multi year:
The file multiYearStatistics.bas prefers a workbook divided into worksheets based on the year of the recorded data. The macro then prints the summary of each year's data to the corresponding sheet.
____________
## run
The link below describes the steps to run a macro in excel very clearly. Just copy-paste all the code inside a module and click run or press F5.
https://www.ablebits.com/office-addins-blog/2013/12/06/add-run-vba-macro-excel/

#### Note
- Make sure to save the excel file as macro-enabled excel workbook
- The maximum number of unique tickers is 10000. If there is more than that, just increase the size of the arrays from the first line.
- See the images (.png) for the multi year data to see an example data and summary.
______________
### Bug report
For any bugs, please email to tolcaglar@gmail.com by giving as much information as possible.
