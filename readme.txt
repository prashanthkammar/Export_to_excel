1) Create a virtualenv env inside ExportToExcel folder
2) Activate virtualenv: env\Scripts\activate.bat
3) pip install -r requirements.txt
4) Login as an admin in flask app and go to User table
5) save (ctrl + s) the User table page as User.html and put that html inside this folder
6) if there are more than 1 pages in User table (1, 2, 3,...) save each page as 1.html, 2.html, 3.html ...
and in terminal run "cat 1.html 2.html 3.html > User.html"
7) Now save the combined User.html inside this folder (replace existing one if data is new)
8) Do the same thing for Track_Users table and save it as TrackUsers.html inside this folder (replace existing one if data is new)
9) run the python file from terminal python html_to_excel.py <table_name>
    ex: python html_to_excel.py user // for User table
        python html_to_excel.py track_users // for Track_Users table
10) There will be a corresponding excel file created for them. The timezone information has been corrected in this code.