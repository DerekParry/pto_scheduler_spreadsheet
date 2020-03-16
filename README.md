## PTO Scheduler
Google Sheet with custom menu item to plan PTO use. 
Used as a workaround to bugs in Paycor. 
Project contains an xlsx file and a Google Apps Script file.

Setup Instructions:
1. Upload the PTO Scheduler spreadsheet to Google Drive
2. Open as a Google Sheet
3. Go to *File => Save as Google Sheets*
4. In the new Sheet, select *Tools => Script editor*
5. Replace the contents of the code file with the contents of *Code.gs* from this project, then save. You can name the project anything you want (e.g. "PTOScheduler").
6. Reload your Google Sheet with the browser's refresh button (refresh or hard refresh via hotkeys in Google Sheets may not fully refresh the page). You will know this step is complete when you see a "PTO Scheduler" menu item in the top bar, next to "Help". This menu item may take a few seconds to load after refreshing.
7. Follow the remaining instructions in the "Instructions" tab of the Sheet

NOTE / REASSURANCE: When you run this script for the first time, you will have to authenticate the application with your Google account. You will need to ignore the Google safety warning about this being an unverified app. This project is owned by, and was entirely set up via these instructions by, your Google account - so you can rest easy that there is no actual threat! Typically, the script isn't actually triggered after the initial authentication process, so you will need to run it once more via *PTO Scheduler => Calculate!*
