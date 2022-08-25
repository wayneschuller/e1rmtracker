# e1rmtracker
# Google App Script to process Barbell Logic client CSV and graph strength progress in Google Sheets

# Input
Currently the code processes CSV data from the Barbell Logic BLOC app. (Dashboard->Workout History)

# Using via Google Sheets
You can cut and paste this code into a Google Sheet via the Extensions->Apps Script menu. Reload the spreadsheet and you should have a strength tracker menu and a welcome sheet with instructions. 

You need to upload your BLOC workout data into a different Google Sheet and paste the URL into the welcome page of the strength tracker. 

You will end up with four graphs for each of: Back Squat, Deadlift, Bench Press, Press. They look something like:
![sheets_strength_sample](https://user-images.githubusercontent.com/1592295/186642113-090e9663-303f-4085-9f28-1d632cae7a1c.jpg)

# Javascript more advanced project
*2020 Aug Update:* Because of limitations with the Google sheets charting system, my development of this project has moved over to a charts.js version that does not rely on Google Drive or Sheets.
https://github.com/wayneschuller/powerlifting_strength_tracker_js

Please contact me if you need help: wayne@schuller.id.au
