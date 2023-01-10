# Google Form to .docx via python
Using python library with gspread and docxtpl for automatic create docx file
from template using google api

This is a prototype

## Todo

- [ ] Create user interface
- [ ] Create user input

## Currently
Can only use only specific device

## How to use 
(Only use before user interface implementation )

1. Go to google api and click create new project
2. Select current project and then click on Credentials
3. Click on manage account then click on CREATE SERVICE ACCOUNT
4. Put detail then click on CREATE AND CONTINUE
5. Click on select role and then go to Basic choose Editor then click CONTINUE
6. Click done no need to insert
7. Click on you created account then go to KEYS
8. Click ADD KEY then create new KEY
9. Choose JSON and then click CREATE
10. After that git clone this repo and put key inside repo folder then change name of the file to "cerds"
11. Then go to your Google Spreadsheet select your sheet that you want to use this program and then click Share
12. Add your service account email link to your Spreadsheet
13. Go to format folder and setup your format
14. Every variable you want to insert must be in double curly branket such as {{ variable }}
15. After that you can run program

<strong>You need to setup only one time and then you can run without any problem </strong>

16. Run "run.bat" and your output will be inside output folder 

## Limitation
The name of variable must not have space in between 
