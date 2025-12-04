# SAP Automation Scripts (VBS)
---------------------------------------------------------------------------------------

# This repository contains .vbs scripts that automate work in SAP.
The main goal of these scripts is to download reports from SAP and save them in a folder chosen by the user.
---------------------------------------------------------------------------------------

# What these scripts do
1. Stable report saving
The scripts make sure that reports are saved in the correct folder every time, without errors.
2. Working with many SAP sessions
Sometimes users are logged into SAP with several sessions open.
The scripts were improved to work correctly even when many SAP windows are open.
3. This includes handling the first window after login:

application.Children(0)
---------------------------------------------------------------------------------------

# Profit
These scripts helped solve common problems:
 * About 60% faster work compared to doing all steps manually.
 * Fewer mistakes because tasks are done the same way every time.
---------------------------------------------------------------------------------------

# Requirements
To use these scripts, you need:
1.	To be logged in to SAP (Login is not automated for security reasons.)
2.	SAP GUI Scripting must be turned on
3.	A Windows system (.vbs runs with Windows Script Host)
4.	SAP GUI that supports scripting (No version number needed)
5.	Your own save path inside the script
In each script there is a comment showing where to put your folder path.
---------------------------------------------------------------------------------------

# How to run the script
1.	Open SAP and log in.
2.	Make sure SAP GUI Scripting is enabled.
3.	Keep only one main SAP window open if possible.
4.	Edit the .vbs file and add your folder path.
5.	Run the script (double-click or use cscript / wscript).


The scripts are stable, but sometimes problems can happen when:
 * many SAP windows are open,
 * SAP shows an extra popup window,
 * another SAP session becomes active.
---------------------------------------------------------------------------------------

# FAQ
1. The script cannot find the SAP window
 * Make sure only one SAP window is open
 * Check that SAP GUI Scripting is turned on
 * Click on the SAP window to make it active before starting the script

2. The script does not save the report
 * The folder path in the .vbs file is correct
 * You have permission to save in that folder
 * The SAP report really creates an output file

3. I see the message “Scripting is disabled”
This means scripting is turned off.
 * Turn on Scripting in SAP GUI settings
 * Restart SAP after changing the setting

4. The script stops in the middle
 * a new popup appears
 * SAP changes the screen layout
 * another SAP window takes focus
Fix: close extra windows and run the script again.
