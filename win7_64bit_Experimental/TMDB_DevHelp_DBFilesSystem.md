# TMDB File System Explained

## Overview of the file system
Originally, the idea was to keep a working directory to test code changes and have directories for creating a blank DB and demo db.
There were also folders for real clubs with actual club data in them.  For privacy reasons, the Toastmaster club folders have been 
removed from this project.  Folder conventions and folders are explained here.  

This documentation is not comprehensive but should shed some light on what is in this project for those interested in using this code.

## Subfolder Conventions

This installation was set up with a root path of c:/ToastMastersDB_x64/
- MS Access uses non-relative paths for linked tables to external files (this is a limitation of MS Access, not the code)
- the design separates data files from the database design in subfolders as explained below
- instructions are in the user help mentioned below for how to relink the tables after installation on your machine if you use a different path  

- Each copy of the database needs the following folders:
  - /data
  - /excel_macros
  - /Reporting  

- Thes folders are customized for each DB as follows
  - for a Toastmasters (TM) club, append a short abbreviation or club number at end of the data folder
  - Examples:
    - data_myClub
	  - all data and data_myClub folders have data tables used by each database
	  - TM-DB-Config:  tables used in database configuration
	  - TM-DB-Data:  Actual data the database is designed to manage (club members, who is in what role, who did what, etc.)
    - excel_macros
	  - file(s) in here are part of the code that generates final reports from the database
    - Reporting
      - this folder is where reporting macros in the database export/write reports as they publish them	

Note:  The Above structure allows keeping of all working folders in one tree.  Each club has its own Data but they will
share the same macro code folder and Reporting output folders.  DB for each club (or data folder) sit in the root at:  
C:\ToastMastersDB_x64
- "x64" - 
  - this version was set up on a 64bit Windows7 machine
  - code was written and edited in Office 2007, then re-saved in Office 2013

## Actual Files and Folders of This Project Explained
- **/demo_converter** - code written to help take a club DB or Working DB and:
  - scrub off usernames and other private info before sharing with public
  - set up Demo DB
  - set up blank "main db" that can be used by others to create a fresh DB
  - status: 
    - this code was under development in 2012
	- current state of it has not been tested
	- code may be incomplete (needs to be checked)  

- **/templates**  
  - tempates for common tasks associated with using the Database System
  - example:  has a worksheet with simple macros to convert emails data:
    - user pastes in email list from the database
	- macro converts it to format that can be copied into To: field of an email  

- Database Files in /ToastMastersDB_x64/
  - _Working:  Working file hooked up to test data (duplicate of demo DB as of this writing)
    - Data files reside /Data_Working
	- Make all changes here and use this fil to generate the rest of the files  

  - _Data:  blank DB that a user can begin inputting content into
    - Data files reside /Data
	- to change code, edit working file and then re-create this file  
 
  - _Demo:  database populated with fake (fun) data using mostly the names of famous and historical figures
    - Data files reside /Data_Demo
	- File can be played with to see how the database works ahead of doing things to live data
	- File can also be used as baseline comparison to _Working to see how live version worked in comparison to new features  

Note: menu systems in all DBs has a "Developers Only" dashboard.  The password for this is the following exact phrase:
- "Mitch says to let me in."
- include the period, do not include the quotes
- As per dev-help, it is recommended to change this password before a live roll-out to users	

- **../help** -
  - in this project the help folder is one level up (from the 64bit code) since it applies to all versions
  - you may want to create a help folder in your distributions and move the content into it
  - what help documentation exists for Database Users and developers goes here
  - these files should get rolled out with code releases to help users  

- **../help_research** -
  - in this project the help folder is one level up since it applies to all versions
  - has folders of research to help with future development  
	
Hope this helps, <br/>
![](https://github.com/TheMitchWorksPro/TestProject/blob/master/html_mitch_logo/Mitch_LogoBG.gif)
