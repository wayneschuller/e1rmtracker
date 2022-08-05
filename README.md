# e1rmtracker
# Google App Script to process and graph strength progress in Google Sheets

# Input
Currently the code processes CSV data from the BLOC app.
Future versions I hope to add support for the BTWB CSV data and custom data.

# Using via Google Sheets
You can cut and paste this code into a Google Sheet via the Extensions->Apps Script menu. From there you can run the code.

# Using via Google clasp (Command Line Apps Script Projects)

(this is mainly for developers)

Assuming you have clasp installed: https://github.com/google/clasp
clasp login  (this will open a browser to authenticate with your google account)

To run the strength tracker, check out the repo and in the same folder run:
clasp create --title "Strength Tracker Sheet"  (choose Sheet from the menu)

clasp push

You will now have a Google Sheet where you can insert your BLOC csv data.
