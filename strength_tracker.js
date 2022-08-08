// 2022 Wayne Schuller

// Globals - here are the BLOC column names
DATEFIELD = "workout_date";
COMPLETED = "workout_completed";
EXERCISENAME = "exercise_name";
ASSIGNEDREPS = "assigned_reps";
ASSIGNEDWEIGHT = "assigned_weight";
ACTUALREPS = "actual_reps";
ACTUALWEIGHT = "actual_weight";
MISSED = "assigned_exercise_missed";


function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Strength Tracker')
        .addItem('Process BLOC CSV data', 'process_BLOC_data')
        .addToUi();
    
    // Check for Welcome sheet.
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Welcome");
    if (!sheet) create_welcome_sheet();
}

// Always have the first sheet as a welcome sheet with instructions and a place to put config URLs
function create_welcome_sheet() {
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
function process_BLOC_data() {
    
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
        let col = 5; // throw all spare lifts into column 5 of the new sheet
        switch (values[row][exercise_name_COL]) {
            case "Squat":
                col = 1;
                break;
            case "Bench Press":
                col = 2;
                break;
            case "Deadlift":
                col = 3;
                break;
            case "Press":
                col = 4;
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
                }
            datefound = true;
            }
        }

        // If this is a new date then create a new row in our collected dataset here.
        if (!datefound) {
            let newrow = [values[row][workout_date_COL], "", "", "", "", ""];
            newrow.fill(onerepmax, col, col+1);
            outputValues.push(newrow);
            bloc_found = true;  // Is this really needed?
        }
    }
  
    // Output the collected data to a special sheet
    let outputsheet = ss.getSheetByName("Processed E1RMs");
    if (outputsheet) ss.deleteSheet(outputsheet);
    outputsheet = ss.insertSheet().setName("Processed E1RMs");

      
    ss.appendRow(["Date","Squat","Bench","Deadlift","Press","Other"]); // Headings
    outputsheet.getRange(1, 1, 1, 6).setFontWeight("bold");

    // Output data into sheet starting in row 2 (row 1 is headings)
    outputsheet.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
        
    // Draw a basic Squat progress chart
    let chartSquatRange = outputsheet.getRange('A:B');

    // Look for existing Squat Progress sheet and delete it
    let squatSheet = ss.getSheetByName("Squat Progress");
    if (squatSheet) ss.deleteSheet(squatSheet);    
      
    let chart = outputsheet.newChart().asLineChart()
          .addRange(chartSquatRange)
          .setTitle('Squat Progress')
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

    outputsheet.insertChart(chart);

    // charts in their own objectsheet look big and beautiful 
    // I have yet to find a way to simply create an objectsheet with a chart
    // The only way seems to be to make a chart then call moveChartToObjectSheet
    squatSheet = ss.moveChartToObjectSheet(chart);
    squatSheet.setName("Squat Progress");
    squatSheet.activate();
    ss.setActiveSheet(squatSheet);
}

// Return a rounded 1 rep max using Epley formula
// For theory see: https://en.wikipedia.org/wiki/One-repetition_maximum 
// Later on we can add different methods
// We really only need a method that works for 1-10 reps.
function estimateE1RM(reps, weight) {
    if (reps === 1) return weight; // If it was a heavy single lets not round it
    return Math.round(weight*(1+reps/30));
}
