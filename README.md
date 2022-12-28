# Application for the automatic improvement of inventory levels

The application downloads all unconverted documents (documents which use the wrong products) between a chosen date range. Create new documents which fix errors. Save changes in Subiekt GT (SQL Server).

<p align="center">
<img src="README_gif.gif">
</p>  

## Technologies
* Python 3.8
  * Object oriented programing
* GUI Programing
  * PyQt6
* SQL
  * SQL Server
  * SQLite
* OLE Automation/ COM technology
* Subiekt GT API


## Ghraphical Explanation of working
<p align="center">
<img src="README_grafic_explain.png">
</p>

<p align="center">
1. The program turns on. Reads the initial data from SQLite (id of all converted documents, and date of the last converted document).
</p>
<p align="center">
2. A user chooses the date range to find all documents which have to be corrected (between this range). Import these documents from SQL Server to my program and start conversion and correct mistakes.
</p>
<p align="center">
3. Creates new documents in Subiekt GT (using COM Microsoft technology) that correct stock level of goods
</p>
<p align="center">
4. Saves in SQLite information about corrected documents (id numbers and date)
</p>

## Summary
### Purpose of application
The application aims to fix system errors. Because one good have more than one instance in the database is causing an increase in the differences in stock levels during the time. Because the number of documents increased by 300% during the last 10 years it is impossible to fix it manually. This application fixes stock levels automatically.
   
### Achivments
1. Saves a employee about 4 hour a week.
2. Makes that the stock levels match, resulting in fewer errors.
3. It allows you to control the prices of goods
### Note
Application's code was presented only to show skils of author.  
The application was design to coperate with two databases (Without them it does not work).   
In code logins and passwords were replace by random string in order to preserve security.
