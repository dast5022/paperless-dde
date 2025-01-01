# paperless-ngx document importer
This is a template for Microsoft Excel to import documents data from paperless-ngx via API to Excel.
# Compatibility
Tested with [paperless-ngx](https://github.com/paperless-ngx/paperless-ngx) Release 2.13.5, Microsoft Excel 2019 on Microsoft Windows 11 and more than 5.000 documents.
# Prerequisites
- you need a licence of Microsoft Excel to use this template
- you need to have access to an instanz of paperless-ngx with a personal token (generate via Django 
- you need to have the rights zu read documents, storage paths, users, document type, correspondents and tags in that instance
# Installation
You have two options to install this template on your own system:
## Short way
You risk downloading an Excel file with macros from Github and running it on your system, then
- download the excel file ending with ".xlsm" (including macros)
- open it and accept using macros, when Excel will ask you
- insert the IP-address of your paperless instance
- insert your personal token, when template is asking you
- this file is configured not to save personal informations like the last person who saved file. You can allows this again by activating on tab "File" of MS Excel main menue
## Long way
You want to play it safe and download the source code yourself, then
- inspect source code here on GitHub
- download the source code file ending with ".bas"
- start MS Excel and open new empty worksheet
- use "save as" to save it with a name you prefer but change file type to ".xlsm" (with macros)
- use tab "Developer" and add a new macro named "makro1" or directly switch to VBA coding editor with menue "View Code"
- in VBA coding editor: import th file eding with ".bas" via "Files"
- change back to normal user interface of MS Excel and start the macro "A_BuildApplication"
- the application is build by itself, you can now use it:
- insert the IP-address of your paperless instance
- insert your personal token, when template is asking you
- this file is configured not to save personal informations like the last person who saved file. You can allows this again by activating on tab "File" of MS Excel 
# Use and Configuration
- You could change the query to select documents from your database. For example find out the id of other tags and fill in into query. Your could also try http://your-instance/api/ an play around with filters and afterwards copy a query from the URL.
- You could remove columns on sheet "documents" as you wish or change their order.
- You can add columns named like your custom fields
- If you have documents with more than 4 tags, you could add more columns with "tag" in first row. They have to be all next to column "tagStart".
# Know issues and open tasks
- bring tags into numeric order
- import tags seperated by komma in one row as alternative to seperated rows
- progress bar or something similar
# Acknowledgments
I used code from an answer of Daniel Ferry ("Excel Hero") on stackoverflow (https://stackoverflow.com/questions/6627652/). His code helps to parse JSON in Excel without using additional libaries. Many thanks to him!

