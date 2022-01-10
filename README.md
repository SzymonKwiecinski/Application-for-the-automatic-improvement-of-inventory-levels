# Application for the automatic improvement of inventory levels

GUI Aplication. Pobiera wszystkie dokumenty z danego przediału czasowego, analizuje je. Tworzy na podstawie analizy nowe dokumenty korygujące błedy. Pozwala zapisać te dokumenty w systemie Subiekt GT 

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
1. The program turns on. Reads the initial data from SQLite (id of all converted documents, and date of last converted document).
</p>
<p align="center">
2. The user chooses the date range to find all documents to correct between this range. Download from SQL Server these documents to my program and start conversion and correct mistakes.
</p>
<p align="center">
3. Creates new documents in Subiekt GT (using COM Microsoft technology) that correct stock level of goods
</p>
<p align="center">
4. Saves in SQLite information about corrected documents (id numbers and date)
</p>

## Summary
### Purpose of application
Celem aplikacji jest naprawienie błedu projektowego sytemu wymiany towarów Kanban. Ze względu na to że jeden towar ma kilka instacji w bazie danych powoduje to po czasie błedy i róznice w stanch faktycznych i magazynowych. 10 lat temu została podjęta decyzja że stany będzie wyrównywał ręcznie jeden z pracowników. Jednak od tego czasu liczba dokumentów zwiększyła się o ponad 100% co sprawia że poprawa manualna błędów jest nie możliwa. Program powstał aby poprawiać te błedy automatycznie.

### Big picture. How it works step by step 
1. Our employee takes from our client's warehouse empty boxes with barcodes. On barcodes are implemented information about the product like Id code, storage location, quantity in the box. 
2. An employee scans barcodes and uploads them to our MS Server database (using another program that I created:). Based on these data, create a warehouse document.
3. In the warehouse, an employee prepares an order and fills boxes with goods.   
4. Then someone sends the warehouse document created earlier to our client using my program. The program does the following:  
   1. Connects to SQL Server database
   2. Exports chosen warehouse document
   3. Converts it to a specific format given by the client 
   4. Saves it as XML file
   5. Connects to Google Cloud Virtual Maschine using SSH authentication and sends a file there
   6. Our client downloads xml file from Google Cloud and easily imports it to its system in this example SAP Cloud platform
### Extra Features:
   1. Ability to connect to Google Cloud Virtual Machine and check files on them
      1. Includes a delete file options
   2. Ability to check histry of all operation and errors 
   
### Achivments
1. Saves a employee about 4 hour a week.
2. Makes that the stock levels match, resulting in fewer errors.
3. It allows you to control the prices of goods
### Note
Application's code was presented only to show skils of author.  
The application was design to coperate with two databases (Without them it does not work).   
In code logins and passwords were replace by random string in order to preserve security.