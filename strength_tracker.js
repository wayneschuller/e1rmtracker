// 2022 Wayne Schuller
//  
// Display Google Sheets e1rm graph comparing different rep + set lift combinations over time
//
// https://github.com/wayneschuller/e1rmtracker
// 
// Released under GPL3 license. https://www.gnu.org/licenses/gpl-3.0.en.html



function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Strength Tracker')
        .addItem('Process BLOC CSV data', 'processBLOCData')
        .addToUi();
    
    // Check for Welcome sheet.
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Welcome");
    if (!sheet) createWelcomeSheet();
}

// Always have the first sheet as a welcome sheet with instructions and a place to put config URLs
function createWelcomeSheet() {
    console.log("Creating welcome sheet");
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('A1').activate();
    spreadsheet.insertSheet(1);
    spreadsheet.getCurrentCell().setValue('Welcome');
    spreadsheet.getRange('A2').activate();
    spreadsheet.getCurrentCell().setValue('Enter full URL for your BLOC CSV file on Google Drive:');
    spreadsheet.getRange('A3').activate();
    spreadsheet.getActiveSheet().setColumnWidth(1, 382);
    spreadsheet.getActiveSheet().setColumnWidth(2, 325);
    spreadsheet.getRange('A1').activate();
    spreadsheet.getActiveRangeList().setFontWeight('bold');
    spreadsheet.getActiveSheet().setName('Welcome');
    spreadsheet.moveActiveSheet(1);
}

// Find the BLOC data and iterate row by row and collect data with top lifts for each session
// FIXME: try to handle any errors gracefully with a dialog explanation
function processBLOCData() {

    // Here are the BLOC column names from their CSV export as of 2022
    const DATEFIELD = "workout_date";
    const COMPLETED = "workout_completed";
    const EXERCISENAME = "exercise_name";
    const ASSIGNEDREPS = "assigned_reps";
    const ASSIGNEDWEIGHT = "assigned_weight";
    const ACTUALREPS = "actual_reps";
    const ACTUALWEIGHT = "actual_weight";
    const MISSED = "assigned_exercise_missed";
    
    let outputValues = []; // we will store our collected data here

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let csv_url = ss.getRange('Welcome!B2').getValue();

    let bloc_csv_ss = SpreadsheetApp.openByUrl(csv_url);
    
    // Get the first sheet
    let sheet = bloc_csv_ss.getSheets()[0];

    let bloc_found = false;

    if (!sheet) return;

    if (sheet.getType() === SpreadsheetApp.SheetType.OBJECT) return;  // Ignore these own-chart sheets

    let values = sheet.getDataRange().getValues(); 
   
    // Iterate through the header row and learn where the columns are
    // We do not assume that BLOC uses consistent column order.
    let workout_date_COL, completed_COL, exercise_name_COL, assigned_reps_COL, assigned_weight_COL, actual_reps_COL, actual_weight_COL, missed_COL;
    for (let col = 0; col < values[0].length; col++) {

        // Logger.log("Title row. Column name is: %s", values[0][col]);
        switch (values[0][col]) {
            case DATEFIELD:
                workout_date_COL = col;
                break;
            case COMPLETED:
                completed_COL = col;
                break;
            case EXERCISENAME:
                exercise_name_COL = col;
                break;
            case ASSIGNEDREPS:
                assigned_reps_COL = col;
                break;
            case ASSIGNEDWEIGHT:
                assigned_weight_COL = col;
                break;  
            case ACTUALREPS:
                actual_reps_COL = col;
                break;
            case ACTUALWEIGHT:
                actual_weight_COL = col;
                break;
            case MISSED:
                missed_COL = col;
        }
    }

    // Give up if we did not find all our expected BLOC data column names
    if (!workout_date_COL ||
        !completed_COL ||
        !exercise_name_COL ||
        !assigned_reps_COL ||
        !assigned_weight_COL ||
        !actual_reps_COL ||
        !actual_weight_COL) 
        return; 
      
    console.log("Excellent BLOC data set header row.");

    /* Iterate CSV rows in reverse - skip row 0 with titles */
    for (let row = values.length - 1 ; row > 0; row--) {
      
        // Give up on this row if there is no date field (should never happen)
        if (!values[row][workout_date_COL]) continue;

        // Give up on this row if it is not a completed workout
        if (!values[row][completed_COL]) continue; 
       
        // Give up on this row if there is no assigned reps 
        // Happens when coach leaves comments in the app
        if (!values[row][assigned_reps_COL]) continue;
      
        // Give up on this row if it was marked as missed TRUE
        if (values[row][missed_COL]) continue;

        // Chose which column of processed data to place results
        let col = 9; // throw all spare lifts into column 5 of the new sheet
        switch (values[row][exercise_name_COL]) {
            case "Squat":
                col = 1;
                break;
            case "Bench Press":
                col = 3;
                break;
            case "Deadlift":
                col = 5;
                break;
            case "Press":
                col = 7;
                break;
        }

        // Record assigned reps and weight.
        var lifted_reps = values[row][assigned_reps_COL];
        var lifted_weight = values[row][assigned_weight_COL];
     
        // Logger.log("Found date: %s, %s reps at %s kg.", values[row][workout_date_COL], lifted_reps, lifted_weight);

        // Override if there is an actual_reps and actual_weight as well
        // This happens when the person lifts different to what was assigned by their coach
        if (values[row][actual_reps_COL] && values[row][actual_weight_COL]) {
            lifted_reps = values[row][actual_reps_COL];
            lifted_weight = values[row][actual_weight_COL];
        }

        // Calculate 1RM for the set
        let onerepmax = estimateE1RM(lifted_reps, lifted_weight);
                        
            
        // Iterate on the collected data so far to see if we have this date already
        let datefound = false;
        for (let j in outputValues) {
            if (outputValues[j][0].toDateString() == values[row][workout_date_COL].toDateString()) {
                if (onerepmax > outputValues[j][col]) {
                    outputValues[j][col] = onerepmax;
                    if (col != 9) outputValues[j][col+1] = `Lift: ${lifted_reps}\@${lifted_weight}`; 
                }
            datefound = true;
            }
        }

        // If this is a new date then create a new row in our collected dataset here.
        if (!datefound) {
            let newrow = [values[row][workout_date_COL], "", "", "", "", "", "", "", "", ""];
            newrow.fill(onerepmax, col, col+1);
            newrow.fill(`Lift: ${lifted_reps}\@${lifted_weight}`, col+1, col+2); // Put in the top set into the notes column (used for chart tooltips)
            outputValues.push(newrow);
            bloc_found = true;  // Is this really needed?
        }
    }
  
    // Output the collected data to a special sheet
    let outputsheet = ss.getSheetByName("Processed E1RMs");
    if (outputsheet) ss.deleteSheet(outputsheet);
    outputsheet = ss.insertSheet().setName("Processed E1RMs");

    ss.appendRow(["Date","Squat", "Squat Notes", "Bench", "Bench Notes", "Deadlift", "Deadlift Notes", "Press", "Press Notes", "Other"]); // Headings
    outputsheet.getRange(1, 1, 1, 10).setFontWeight("bold");

    // Output data into sheet starting in row 2 (row 1 is headings)
    outputsheet.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
        
    // Draw progress charts
    let dateRange = outputsheet.getRange('A1:A1000');
    let liftRange = outputsheet.getRange('B1:B1000');
    buildE1RMChart('Squat', dateRange, liftRange);

    liftRange = outputsheet.getRange('D1:D1000');
    buildE1RMChart('Bench Press', dateRange, liftRange);
    
    liftRange = outputsheet.getRange('F1:F1000');
    buildE1RMChart('Deadlift', dateRange, liftRange);

    liftRange = outputsheet.getRange('H1:H1000');
    buildE1RMChart('Press', dateRange, liftRange);

}


// Build a new google sheets chart and put in a standalone object sheet 
function buildE1RMChart (exerciseName, dateRange, liftRange) {

    // We need a sheet to build the chart in, so take the active one.
    // It doesn't matter which sheet because we will move the chart to a new sheet
    // All due to a strange way that the GSheets API service works.
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getActiveSheet();
     
    let title = exerciseName + ' Progress';
    
    let chart = sheet.newChart().asLineChart()
          .addRange(dateRange)
          .addRange(liftRange)
          .setTitle(title)
          .setXAxisTitle('Date')
          .setYAxisTitle('Epley One Rep Max')
          .setOption('interpolateNulls', true)
          .setPosition(1, 8, 0, 0)
          .setOption('height', 600)
          .setOption('width', 1000)
          .setNumHeaders(1)
          .setCurveStyle(Charts.CurveStyle.SMOOTH)
          .setPointStyle(Charts.PointStyle.TINY)
          .setColors(["red"])
          .setOption('pointShape', "diamond")
          .setOption('series.0.hasAnnotations', true)
          .setOption('series.0.dataLabel', 'value')
          .build(); 
     
    sheet.insertChart(chart);
    
    // Delete any existing object sheet for this lift
    sheet = spreadsheet.getSheetByName(title);
    if (sheet) spreadsheet.deleteSheet(sheet); 

    // Now we move the chart from it's temp sheet to a new big home
    let newSheet = spreadsheet.moveChartToObjectSheet(chart);
    newSheet.setName(title);
}

// Return a rounded 1 rep max using Epley formula
// For theory see: https://en.wikipedia.org/wiki/One-repetition_maximum 
// Later on we can add different methods
// We really only need a method that works for 1-10 reps.
function estimateE1RM(reps, weight) {
    if (reps === 1) return weight; // If it was a heavy single lets not round it
    return Math.round(weight*(1+reps/30));
}
