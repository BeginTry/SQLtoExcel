# SQL to Excel
Iterates though a folder of *.sql script files, executing each script, capturing the script query output, and writing the results to a Microsoft Excel Worksheet file (one worksheet per script file).

<b>Quick Start:</b><br/>
Download the files from the <a href="https://github.com/BeginTry/SQLtoExcel/releases">latest release</a> and copy one or more SQL scripts into the same folder. Edit the .bat file to specify the name of the SQL instance you want to connect to and the name of the database (db name may or may not be applicable, depending on how the scripts are written). Launch the .bat file and SQLtoExcel.exe will iterate through the SQL scripts in the directory, run each script, capture the script query output, and write it to an Excel spreadsheet. A separate worksheet for each script is created, with the script file name used for the worksheet name.

<b>Other Notes</b><br/>
SQLtoExcel.exe attempts to ignore “GO” batch separators. If there are #temp tables that span batches, you may encounter errors.
<br/><br/>
Blog Post with some screen grabs: <a href="http://itsalljustelectrons.blogspot.com/2018/11/Introducing-SQL-to-Excel.html">Introducing SQL-to-Excel</a>
