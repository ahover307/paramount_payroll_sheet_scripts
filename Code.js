// custom menu
function onOpen() {
    //create menu and sidebar
    SpreadsheetApp.getUi()
        .createMenu('Alex Made this Menu')
        .addItem('Company Selection', 'processOutput').addToUi();
}

//show sidebar
function processOutput() {
    let daySheets = [
        'Monday W1', 'Monday W2',
        'Tuesday W1', 'Tuesday W2',
        'Wednesday W1', 'Wednesday W2',
        'Thursday W1', 'Thursday W2',
        'Friday W1', 'Friday W2'
    ];

    daySheets.forEach(sheetName => {
        let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        let rows = sheet.getRange("A2:" + sheet.getLastColumn() + sheet.getLastRow()).getValues();

        rows.forEach(row => {
            let name = row[0];

            
        })
    })
}