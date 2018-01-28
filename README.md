# MS_Access_TMDB

## What Is This Project?
This project originated as a database for managing data for [Toastmasters clubs](https://www.toastmasters.org/).  It was made available to Toastmasters clubs in 2008 via a [Google project site](https://sites.google.com/site/tmdbtoydb/Home/tm-database-project).  The code was deliberately locked to prevent accidents and the first version was created in Office 97 .mdb files.  Copies of this code base later became a starting point for multiple projects created behind the firewall at work that unfortunately will never see the light of day again.  For a quick visual demo of the Database, see [this presentation](https://github.com/TheMitchWorksPro/TestProject/blob/master/TMDB – MS Access DB - About DB Presentation.pdf).

The code is now being opened up and presented here with the following intentions:

1. Usage of the Database (as it is):
  - manage the information for a Toastmasters clubs
  - if a club has roles and speeches or presentations, it could be configured for any club  

2. Usage of the code:
  - think of this as a "toy database" that is good starter code for concepts that can be applied to other databases
  - this is how I used it in the past

3. What to look for in the DB as a coder
  - Configurable dashboard code for a clean reusable UI that could be applied to any MS Access DB
    - Code built into the dashboard for:
      - "public" dashboards all users see to utilize features
	  - "private" dashboards used only by developers:
	    - a "Developers" dashboard that requires a password for navigation to take you there
	    - a hidden dashboard that can be used to test feature buttons before taking them live
		
  - VB driven reporting that exports queries to an Excel template and hands off to Excel to format the Reports
    - This system uses Excel as the medium for all reports instead of the report features within MS Access
    - Code exports MS Acces SQL queries into a copy of an Excel Template
	- The code then triggers Excel to open the new file and run macros in it to complete report generation
	- The code uses simple strategies to let the user know when the report is done and gives the user the option to:
	  - view the report in Excel
	  - close Excel to open/view the report from it later and continue working in Access

## Code Versions Available
The creator of this project only has access to the most recent version of MS Access / Excel used to generate this code.  Though older
versions are provided, the onus is on the developer/user to debug, enhance, etc. any older version.  Regarding the current version,
testing was conducted to ensure it works on the current system described, but some of the MS Office VB is finicky.  Experience has shown
that particularly the code that communicates between Access and Excel had to be debugged for every change of version and/or hardware 
that was used to run it.  Even the same release of MS Office, when running on VDI (Virtual Desktop Infrastructure) at work, required tweaks 
to the code to make it work right that then did not work on a standard (non-VDI) laptop, resulting in different versions of the code at work
and home.  The versions provided here were undertaken in my spare time for my Toastmasters club.  More advanced implementations of these coding
principles that were implemented at work were left at work as per company policy and, unfortunatley, are not available here.

### Subproject folders:
- [win_Pre7_MSOffice97](win_Pre7_MSOffice97): &nbsp;&nbsp;&nbsp;&nbsp; oldest Office 97 version of the code 
  - Code was tested years ago and ran on a 32bit older Windows Machine (Win95, NT, or XT)
  - The code file would need to be re-linked to the database files in /data or /demo folders
  - an HTML file that links to [this site](https://sites.google.com/site/tmdbtoydb/Home/tm-database-project/tm-database-project-pg2) is provided in this folder;  additional files need to be downloaded from here to complete the distribution.
- [win7_32Bit_Office2003](win7_32Bit_Office2003): &nbsp;&nbsp;&nbsp;&nbsp; version used for two different Toastmasters clubs as-of 2012
  - Code was tested and ran in MS Office 2003 on a 32bit Windows 7 Home Edition laptop
  - Code is expected to also work in MS Office 2007 but this has not been tested
- [win7_64bit_Experimental](win7_64bit_Experimental): &nbsp;&nbsp;&nbsp;&nbsp; Experimental version 
    - Code tested and debugged on a 64bit MS Office 2013 laptop running Windows 7 Home Premium 
	- Hardware includes 16 Gigs Ram and Flash drives
	- Code includes experiments to enhance the button dashboards to support more buttons arranged in 2-columns
	- Reporting code tested and debugged but this version was never used in a live setting
	- /Reporting folder has sample reports from testing using the /demo data 

This [Google Project site](https://sites.google.com/site/tmdbtoydb/Home) has copies of the distributions that were made available to the public from 2008 to 2012 with instructions and help
intended to help with installation of the distributions.  Hope you find this project useful.	
	
![Mitch](https://github.com/TheMitchWorksPro/TestProject/blob/master/html_mitch_logo/Mitch_LogoBG.gif)

