# e1rmtracker
# Google App Script to process Barbell Logic client CSV and graph strength progress in Google Sheets

# Input
Currently the code processes CSV data from the Barbell Logic BLOC app. (Dashboard->Workout History)

Future versions I hope to add support for the Beyond the White Board CSV data and also custom data.

I am also developing a charts.js version that does not rely on Google Drive or Sheets.

# Using via Google Sheets
You can cut and paste this code into a Google Sheet via the Extensions->Apps Script menu. Reload the spreadsheet and you should have a strength tracker menu and a welcome sheet with instructions. 

You need to upload your BLOC workout data into a different Google Sheet and paste the URL into the welcome page of the strength tracker. 
