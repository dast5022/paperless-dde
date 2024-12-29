# paperless-ngx document importer
This is a template for Microsoft Excel to import documents data from paperless-ngx via API to Excel.
# Compatibility
Tested with [paperless-ngx](https://github.com/paperless-ngx/paperless-ngx) Release 2.13.5, Microsoft Excel 2019 on Microsoft Windows 11 and more than 5.000 documents.
# Prerequisites
- you need a licence of Microsoft Excel to use this template
- you need to have access to an instanz of paperless-ngx with a personal token (generate via Django 
- you need to have the rights zu read documents, storage paths, users, document type, correspondents and tags in that instance
# Installation
- download this excel file
- accept using macros, when Excel will ask you
- insert the IP-address of your paperless instance
- insert your personal token, when template is asking you
- this file is configured not to save personal informations like the last person who saved file. You can allows this again by activating on tab "File" of MS Excel main menue
# Use and Configuration
- You could change the query to select documents from your database. For example find out the id of other tags and fill in into query. Your could also try http://your-instance/api/ an play around with filters and afterwards copy a query from the URL.
- You could remove columns on sheet "documents" as you wish or change their order.
- If you have documents with more than 4 tags, you could add more columns with "tag" in first row. They have to be all next to column "tagStart".
# Know issues and open tasks
- importing notes
- importing custom fields
- publishing version with self-building code
- bring tags into numeric order
- import tags seperated by komma in one row as alternative to seperated rows
