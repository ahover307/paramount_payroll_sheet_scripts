// custom menu
function onOpen() {
    //create menu and sidebar
    SpreadsheetApp.getUi()
        .createMenu('Alex Made this Menu')
        .addItem('Company Selection', 'processOutput').addToUi();
}

const lengthOfTemplate = 40;

//show sidebar
function processOutput() {
    let daySheets = [
        'Monday',
        'Tuesday',
        'Wednesday',
        'Thursday',
        'Friday',
        'Extra'
    ];

    let outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Output');
    outputSheet.clear()
    let outputRowTracker = 1;

    let templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template")
    let templateHeights = [];
    let templateRangeW1 = templateSheet.getRange("A1:D40")
    let templateRangeW2 = templateSheet.getRange("A41:D80")
    let total = 0;
    for (let i = 1; i < 41; i++) {
        templateHeights.push(templateSheet.getRowHeight(i));
        let temp = templateSheet.getRowHeight(i)
        total += temp
    }
    for (let i = 1; i <= 4; i++) {
        outputSheet.setColumnWidth(i, templateSheet.getColumnWidth(i));
    }

    daySheets.forEach(sheetName => {
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        let rows = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();

        rows.forEach(row => {
            let name = row[0];
            let times = []
            for (let i = 1; i < row.length; i++) {
                if (row[i] !== '') {
                    times.push(row[i]);
                }
            }

            //    Copy template, and replace data on each row.
            let templateData = templateRangeW1.getValues();
            for (let i = 0; i < templateData.length; i++) {
                for (let j = 0; j < templateData[i].length; j++) {
                    if (templateData[i][j] === '{NAME}') {
                        templateData[i][j] = name;
                    } else if (templateData[i][j] === '{TIME}') {
                        templateData[i][j] = times.pop();
                    }
                }
            }

            outputSheet.getRange("A" + outputRowTracker + ":D" + (outputRowTracker + lengthOfTemplate - 1))
                .setValues(templateData);

            for (let i = 'A'; i <= 'D'; i++) {
                for (let j = 1; j < 41; j++) {
                    outputSheet.setRowHeight((j + outputRowTracker - 1), templateHeights[j])
                    outputSheet.getRange(i + (j + outputRowTracker - 1))
                        .setFontWeight(templateSheet.getRange(i + j).getFontWeight());
                    outputSheet.getRange(i + (j + outputRowTracker - 1))
                        .setFontColor(templateSheet.getRange(i + j).getFontColor());
                    outputSheet.getRange(i + (j + outputRowTracker - 1))
                        .setFontSize(templateSheet.getRange(i + j).getFontSize());
                    outputSheet.getRange(i + (j + outputRowTracker - 1))
                        .setHorizontalAlignment(templateSheet.getRange(i + j).getHorizontalAlignment());
                    outputSheet.getRange(i + (j + outputRowTracker - 1))
                        .setVerticalAlignment(templateSheet.getRange(i + j).getVerticalAlignment());


                }
            }

            //    Paste onto the output sheet

            //
        })
    })
}