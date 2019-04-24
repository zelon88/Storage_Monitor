NAME: Storage_Monitor.vbs

TYPE: Visual Basic Script

PRIMARY LANGUAGE: VBS

AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 5/31/2018

CURRENT VERSION DATE: 4/24/2019

VERSION: v2.1 - Fix more bugs with running as a SYSTEM task.

DESCRIPTION: 
A simple script for monitoring storage devices for added/removed volumes and and adequate disk space.

PURPOSE: 
To monitor company storage devices for changes and/or disk space issues that need to be manually addressed.

INSTALLATION INSTRUCTIONS: 
1. Copy the entire "Storage_Monitor" folder into the "AutomationScripts" folder on SERVER (or any other network accesbible location).
2. Fill out the sendmail.ini file with your email server address and credentials.
3. Add a scheduled task to run "Storage_Monitor.vbs" every 30 minutes.
4. Ensure that everyone who runs the script can modify the contents of "Warning.mail" in the AutomationScripts folder and execute sendmail.exe.

NOTES: SendMail for Windows is required and included in the "Storage_Monitor" folder. The SendMail data files must be included in the same directory as "Data_Monitor.vbs" in order for emails to be sent correctly.
