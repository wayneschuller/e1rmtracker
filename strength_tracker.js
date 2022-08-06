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
}


// Find the BLOC data and iterate row by row and collect data with top lifts for each session
function process_BLOC_data() {
    
    let outputValues = []; // we will store our collected data here
    let ss = SpreadsheetApp.getActiveSpreadsheet(); // FIXME: whats the cool way to check for null etc

    // Loop through all the sheets in the spreadsheet looking for the BLOC data 
    let allsheets = ss.getSheets();

    let bloc_found = false;

    for (let s in allsheets) {
        let sheet = allsheets[s];
        let values = sheet.getDataRange().getValues();
    
        // Dynamically find where the columns are
        // We do not assume that BLOC uses consistent column order.
        let workout_date_COL, completed_COL, exercise_name_COL, assigned_reps_COL, assigned_weight_COL, actual_reps_COL, actual_weight_COL, missed_COL;

        for (let col = 0; col < values[0].length; col++) {

            Logger.log("Title row. Column name is: %s", values[0][col]);
                
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
    
            // Give up on this sheet if we did not find all our expected BLOC data column names
            if (!workout_date_COL ||
                !completed_COL ||
                !exercise_name_COL ||
                !assigned_reps_COL ||
                !assigned_weight_COL ||
                !actual_reps_COL ||
                !actual_weight_COL) 
            continue; 

            /* Iterate backwards in chronological order - skip row 0 with titles */
            for (let row = values.length - 1 ; row > 1; row--) {
      
            // Give up on this row if there is no date field (should never happen)
            if (!values[row][workout_date_COL]) continue;

            // Give up on this row if it is not a completed workout
            if (!values[row][completed_COL]) continue; 
       
            // Give up on this row if there is no assigned reps 
            // Normally this is just coach programming comments in the BLOC data
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

            // Calculate 1RM for the list lift
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
            }
            }
  
        // Output the collected data to a new sheet
        let outputsheet = ss.insertSheet(); 

        ss.appendRow(["Date","Squat","Bench","Deadlift","Press","Other"]); // Headings
        outputsheet.getRange(1, 1, 1, 6).setFontWeight("bold");

        // Output data into sheet starting in row 2 (row 1 is headings)
        outputsheet.getRange(2, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
        
        // Draw a basic Squat progress chart
        // FIXME: we need to turn on the 'plot null values' option to join the dots
        let chartSquatRange = outputsheet.getRange('A:B');
        let lineChartBuilder = outputsheet.newChart().asLineChart();
        let chart = lineChartBuilder
          .addRange(chartSquatRange)
          .setTitle('Squat Progress')
          .setPosition(1, 8, 0, 0)
          .setNumHeaders(1)
          .build(); 

        outputsheet.insertChart(chart);         

        bloc_found = true;
        }
    }

    if (bloc_found) {
        // FIXME: tell the user the sheet name via the ActiveSpreadsheet?
        SpreadsheetApp.getUi().alert('Processed BLOC data into new active sheet.'); 
    } else {
        SpreadsheetApp.getUi().alert('Did not find BLOC data in any sheets. (be sure to do File->Import the BLOC workout data csv into a sheet of this spreadsheet)');
    }
}

// Return a 1 rep max using Epley formula
// For theory see: https://en.wikipedia.org/wiki/One-repetition_maximum 
// Later on we can add different methods
// We really only need a method that works for 1-10 reps.
function estimateE1RM(reps, weight) {
    return weight*(1+reps/30);
}
