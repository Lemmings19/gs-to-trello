// Trello import script modified to fit the Project Roadmap Template provided by Steve after our company training.

// Set up the Google Spreadsheets dropdown menu
function onOpen(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Display Available Members", functionName: "displayMembers"},{name: "Display Available Boards", functionName: "displayBoards"},{name: "Display Available Lists", functionName: "displayLists"},{name: "Upload Outstanding Backlog Items", functionName: "upload"}];
    ss.addMenu("Trello", menuEntries);
}

// Get/update the control values.
function checkControlValues(requireList, requireBoard) {
    var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UploadControls").getRange("B3:B7").getValues();

    var appKey = col[0][0].toString().trim();
    if (appKey == "") {
        return "App Key not found";
    }
    ScriptProperties.setProperty("appKey", appKey);

    var token = col[1][0].toString().trim();
    if (token == "") {
        return "Token not found";
    }
    ScriptProperties.setProperty("token", token);

    if (requireBoard) {
        var bid = col[2][0].toString().trim();

        if(bid == "") {
            return "Board ID not found";
        }
        ScriptProperties.setProperty("boardID", bid);
    }

    if (requireList) {
        var lid = col[3][0].toString().trim();
        ScriptProperties.setProperty("listID", lid);
    }

    var sheetName = col[4][0].toString().trim();
    if (sheetName == "") {
        return "No sheet selected.";
    } else if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
        return "Sheet not found.";
    }
    ScriptProperties.setProperty("sheetName", sheetName);

    return "";
}

// Commit spreadsheet cells to Trello
function upload() {
    var statusCol = 0;
    var clientCol = 1;
    var epicCol = 2;
    var titleCol = 4;
    var commentCol = 5;
    var mvpCol = 6;
    var cycleCol = 7;
    var dueDateCol = 8;
    var dollarValueCol = 10;
    var hoursCol = 11;

    var startRow = 10;

    var startTime = new Date();
    Logger.log("Started at:" + startTime);
    var error = checkControlValues(true, true);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ScriptProperties.getProperty("sheetName"));
    var existingLabels = getExistingLabels(ScriptProperties.getProperty("boardID"));

    if (existingLabels == null || existingLabels.length == 0) {
        return;
    }

    var successCount = 0;
    var partialCount = 0;
    var rows = sheet.getDataRange().getValues();

    for (var i = startRow ; i < rows.length ; i++) {
        var currentRow = rows[i];

        // Only process if there is a card title and an epic.
        if (currentRow[titleCol].trim() != "" && currentRow[epicCol] != "") {
            r = i + 1;

            var status = currentRow[statusCol];

            currentTime = new Date();

            Logger.log("Row " + r + ":" + currentTime);

            if (currentTime.valueOf() - startTime.valueOf() >= 330000) { // 5.5 minutes - scripts time out at 6 minutes
                Browser.msgBox("NOTICE: Script was about to time out so upload has been terminated gracefully ." + successCount + " backlog items were uploaded successfully.");
                return;
            } else if (status == ".") { // Row already processed.
                Logger.log("Ignoring row " + r + ". Status column indicates already imported.");
            } else if (status == "x") {
                Browser.msgBox("ERROR: Row " + r + " indicates that it was partially created the last time this script was run. Verify the card in Trello. If the card is unsatisfactory, clear column A to re-import (includes use cases). Set column A in row " + r + " to '.' to skip it. Ending script.");
                return;
            } else if (status == "") { // Status cell empty. Import row.

                var statusCell = sheet.getRange(r, statusCol + 1, 1, 1);
                var dueDate = null;
                var card = true;

                // Get due date.
                if (currentRow[dueDateCol] !== '') {
                    dueDate = currentRow[dueDateCol];
                }

                // Indicate that this row has begun importing.
                statusCell.setValue("x");

                partialCount++;

                var description = currentRow[commentCol];
                var descriptiveRowCount = i;
                var descriptiveRow;
                var descriptiveStatusCell;

                // Fill in description with use cases that follow it.
                while (card) {
                    descriptiveRowCount++;
                    descriptiveRow = rows[descriptiveRowCount];
                    if (descriptiveRow[titleCol] != "" && descriptiveRow[epicCol] == "") {
                        // Add use case to card description.
                        if (description == "") {
                            description = descriptiveRow[titleCol];
                        } else {
                            description += "\n" + descriptiveRow[titleCol];
                        }
                        // Indicate that the use case has been imported.
                        descriptiveStatusCell = sheet.getRange(descriptiveRowCount + 1, statusCol + 1, 1, 1);
                        descriptiveStatusCell.setValue(".");
                    } else {
                        // Row is empty or is actually a story.
                        // Only stories have epics.
                        break;
                    }

                    if (descriptiveRow > rows.length) {
                        // Prevent infinite loop.
                        break;
                    }
                }
                if (description == "") {
                    description = "Original total hour estimate was " + currentRow[hoursCol] + ".";
                } else {
                    description += "\nOriginal total hour estimate was " + currentRow[hoursCol] + ".";
                }

                // Get card title, description, point estimate, id, due date, assignees
                var card = createTrelloCard(currentRow[titleCol], description, currentRow[hoursCol], ScriptProperties.getProperty("listID"), dueDate, "");

                addTrelloLabels(card.id, currentRow[2], existingLabels);

                // Indicate that this row has been imported.
                statusCell.setValue(".");

                SpreadsheetApp.flush();
                partialCount --;
                successCount ++;
            }
        }
    }

    Browser.msgBox( successCount + " backlog items were uploaded successfully.");

    return;
}

// Displays id's for members which exist in your Trello account.
function displayMembers() {

    var error = checkControlValues(false,true);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var url = constructTrelloURL("boards/"+ ScriptProperties.getProperty("boardID") + "/members");
    var resp = UrlFetchApp.fetch(url, {"method": "get"});
    var values = Utilities.jsonParse(resp.getContentText())

    var app = UiApp.createApplication();

    var header1 = app.createHTML("<b>Member Name</b>");
    var header2 = app.createHTML("<b>Member Id</b>");
    var grid = app.createGrid(values.length+1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i=values.length-1; i>=0; i--) {
        grid.setText(i+1, 0, values[i].fullName);
        grid.setText(i+1, 1, values[i].id);
    }

    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Members");

    SpreadsheetApp.getActiveSpreadsheet().show(app);

    return;
}

// Displays id's for boards which exist in your Trello account.
function displayBoards() {

    var error = checkControlValues(false,false);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var url = constructTrelloURL("members/me/boards");
    var resp = UrlFetchApp.fetch(url, {"method": "get"});
    var values = Utilities.jsonParse(resp.getContentText())

    var app = UiApp.createApplication();

    var header1 = app.createHTML("<b>Board Name</b>");
    var header2 = app.createHTML("<b>Board Id</b>");
    var grid = app.createGrid(values.length+1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i=values.length-1; i>=0; i--) {
        grid.setText(i+1, 0, values[i].name);
        grid.setText(i+1, 1, values[i].id);
    }
    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Boards");

    SpreadsheetApp.getActiveSpreadsheet().show(app);

    return;
}

// Displays id's for checklists which exist in your Trello board.
function displayLists() {

    var error = checkControlValues(false,true);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var url = constructTrelloURL("boards/"+ ScriptProperties.getProperty("boardID") + "/lists");
    var resp = UrlFetchApp.fetch(url, {"method": "get"});
    var values = Utilities.jsonParse(resp.getContentText())

    var app = UiApp.createApplication();

    var header1 = app.createHTML("<b>List Name</b>");
    var header2 = app.createHTML("<b>List Id</b>");
    var grid = app.createGrid(values.length+1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i=values.length-1; i>=0; i--) {
        grid.setText(i+1, 0, values[i].name);
        grid.setText(i+1, 1, values[i].id);
    }

    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Lists");

    SpreadsheetApp.getActiveSpreadsheet().show(app);

    return;
}

function getExistingLabels(boardId) {

    var values = null;
    var url = constructTrelloURL("boards/" + boardId + "/labels");
    var resp = UrlFetchApp.fetch(url, {"method": "get","muteHttpExceptions":true });

    if (resp.getResponseCode() == 200) {
        var values = Utilities.jsonParse(resp.getContentText());
    } else {
        Browser.msgBox("ERROR:Unable to return existing labels from board:" + resp.getContentText());
    }

    return values;
}

function addChecklist(card, boardID, checklistName, checklistData) {

    var data = checklistData.split("\n");
    var checklist = null;

    for (var i = 0; i < data.length; i++) {
        if (data[i] != "") {
            if (checklist == null) {
                checklist = createTrelloChecklist(card.id,checklistName);
            }
            createTrelloChecklistItem(checklist.id,data[i]);
        }

    }

    if (checklist !== null) {
        addTrelloChecklistToCard(checklist.id, card.id);
    }
}

function createTrelloCard(cardName, cardDesc, storyPoints, listID, dueDate,members) {
    var name = cardName;
    
    if (storyPoints != "") {
        name = "(" + storyPoints + ") " + cardName;
    }
    
    var url = constructTrelloURL("cards") + "&name=" + encodeURIComponent(name) + "&desc=" + encodeURIComponent(cardDesc) + "&idList=" + listID + "&due=" + encodeURIComponent(dueDate);

    if (members !="") {
        url += "&idMembers=" + encodeURIComponent(members);
    }
    
    var resp = UrlFetchApp.fetch(url, {"method": "post"});

    return Utilities.jsonParse(resp.getContentText());

}

function createTrelloChecklist(cardID, name) {
    var url = constructTrelloURL("checklists") + "&name=" + encodeURIComponent(name) + "&idCard="  + encodeURIComponent(cardID);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});
    return Utilities.jsonParse(resp.getContentText());
}

function createTrelloComment(cardID, name) {
    var url = constructTrelloURL("cards/"+ cardID + "/actions/comments") + "&text=" + encodeURIComponent(name);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});
    return Utilities.jsonParse(resp.getContentText());
}

function createTrelloAttachment(cardID, attachment) {
    if (attachment == "") {
        return;
    }

    var attachments = attachment.split(",");

    for (var i = 0; i < attachments.length; i++) {
        var url = constructTrelloURL("cards/"+ cardID + "/attachments") + "&url=" + encodeURIComponent(attachments[i]);
        var resp = UrlFetchApp.fetch(url, {"method": "post"});
    }

    return;
}

function addTrelloMember(cardID, member) {
    if (member == "") {
        return;
    }

    var members = member.split(",");

    for (var i = 0; i < members.length; i++) {
        var url = constructTrelloURL("cards/"+ cardID + "/idMembers") + "&value=" + encodeURIComponent(members[i]);
        var resp = UrlFetchApp.fetch(url, {"method": "post"});
    }
    return;
}

function addTrelloLabels(cardID, label, existingLabels) {
    if (label == "" ) {
        return;
    }
    var labels = label.split(",");

    for (var i = 0; i < labels.length; i++) {
        var labelId = getIdForLabelName(labels[i],existingLabels);

        if (labelId == null) {
            var url = constructTrelloURL("cards/"+ cardID + "/labels") + "&color=null&name=" + encodeURIComponent(labels[i]);
            var resp = UrlFetchApp.fetch(url, {"method": "post"});
        } else {
            var url = constructTrelloURL("cards/"+ cardID + "/idLabels") + "&value=" + encodeURIComponent(labelId);
            var resp = UrlFetchApp.fetch(url, {"method": "post"});

        }
    }

    return;
}

function getIdForLabelName(label, existingLabels) {

    for (var i = 0; i < existingLabels.length; i++) {
        if (existingLabels[i].name.toUpperCase() == label.toUpperCase()) {
            return existingLabels[i].id;
        }
    }

    return null;
}

function createTrelloChecklistItem(checkListID, name) {
    var url = constructTrelloURL("checklists/" + checkListID + "/checkItems") + "&name=" + encodeURIComponent(name);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});
    return Utilities.jsonParse(resp.getContentText());
}

function addTrelloChecklistToCard(checkListID, cardID) {
    var url = constructTrelloURL("cards/" + cardID + "/checklists") + "&value=" + encodeURIComponent(checkListID);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});

    return Utilities.jsonParse(resp.getContentText());
}

function constructTrelloURL(baseURL) {
    return "https://trello.com/1/"+ baseURL +"?key="+ScriptProperties.getProperty("appKey")+"&token="+ScriptProperties.getProperty("token");
}
 