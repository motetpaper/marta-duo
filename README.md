# marta-duo is a Google Sheets editor add-on that exports Markdown tables

There are two files in this duo:
  * `Code.gs`: provides the server-side logic, building a Marta object
  * `Sidebar.html`: provides the client-side processing and download of the markdown text file

The purpose of this Google Sheet Add-on is performance and simplicity.

Unlike many other Markdown table exporters for Google Sheets, 
marta-duo will not timeout for very large tables. 

To this point, the first test spreadsheet ([test1.xlsx][testsheets]) has six (6) columns 
and one-thousand (1000) rows.

While it may take a while on some machines, the entire markdown table does indeed export (on tested machines).

Performance metrics and unit tests follow.


[testsheets]: https://github.com/motetpaper/marta-duo/tree/main/tests/test-sheets
