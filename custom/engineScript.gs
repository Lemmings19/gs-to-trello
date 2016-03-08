// Set up the Google Spreadsheets dropdown menu
function onOpen(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Add stories to Trello", functionName: "upload"},
        {name: "Display available boards", functionName: "displayBoards"},
        {name: "Display available lists", functionName: "displayLists"},
        {name: "Display available members", functionName: "displayMembers"}];
    ss.addMenu("Trello", menuEntries);
}

// Get/update the control values.
function checkControlValues(requireList, requireBoard, requireSheet) {
    var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls").getRange("C1:C32").getValues();

    var appKey = col[4][0].toString().trim();
    if (appKey == "") {
        return "App Key not found. Update Controls sheet.";
    }
    ScriptProperties.setProperty("appKey", appKey);

    var token = col[6][0].toString().trim();
    if (token == "") {
        return "Token not found. Update Controls sheet.";
    }
    ScriptProperties.setProperty("token", token);

    if (requireBoard) {
        var bid = col[11][0].toString().trim();

        if(bid == "") {
            return "Board ID not found. Update Controls sheet.";
        }
        ScriptProperties.setProperty("boardID", bid);
    }

    if (requireList) {
        var lid = col[13][0].toString().trim();
        
        if (lid == "") {
            return "List ID not found. Update Controls sheet.";
        }
        ScriptProperties.setProperty("listID", lid);
    }

    if (requireSheet) {
        var sheetName = col[30][0].toString().trim();
        if (sheetName == "") {
            return "No sheet selected. Update Controls sheet.";
        } else if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
            return "Sheet not found. Update Controls sheet.";
        }
        ScriptProperties.setProperty("sheetName", sheetName);
    }

    return "";
}

// Commit spreadsheet cells to Trello
function upload() {
    var statusCol = 1;
    var clientCol = statusCol + 1;
    var epicCol = statusCol + 2;
    var titleCol = statusCol + 4;
    var commentCol = statusCol + 5;
    var mvpCol = statusCol + 6;
    var cycleCol = statusCol + 7;
    var dueDateCol = statusCol + 8;
    var dollarValueCol = statusCol + 10;
    var hoursCol = statusCol + 11;
    var roleCol = statusCol + 13;

    var headerRow = 8;
    var startRow = 10;

    var roleCount = 50;

    var startTime = new Date();
    Logger.log("Started at:" + startTime);
    var error = checkControlValues(true, true, true);
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

            var titles = sheet.getRange(headerRow, roleCol, 1, roleCount).getValues();
            var hours = sheet.getRange(r, roleCol, 1, roleCount).getValues();

            currentTime = new Date();

            Logger.log("Row " + r + ":" + currentTime);

            if (currentTime.valueOf() - startTime.valueOf() >= 330000) { // 5.5 minutes - scripts time out at 6 minutes
                Browser.msgBox("NOTICE: Script was about to time out so upload has been terminated gracefully ." + successCount + " backlog items were uploaded successfully.");
                return;
            } else if (status == ".") { // Row already processed.
                Logger.log("Ignoring row " + r + ". Status column indicates already imported.");
            } else if (status == "x") {
                Browser.msgBox("ERROR: Row " + r + " did not finish creation. Verify the card in Trello. Clear column B to re-import. Set column B to '.' to skip.");
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

                var description = currentRow[commentCol] + "\n\n" + "**Yellow Cards**\n\n";
                var descriptiveRowCount = i;

                // Fill in description with use cases that follow it.
                while (card) {
                    descriptiveRowCount++;
                    var descriptiveRow = rows[descriptiveRowCount];
                    if (descriptiveRow[titleCol] != "" && descriptiveRow[epicCol] == "") {
                        // Add use case to card description.

                        // If multiple lines for a usecase, split it into multiple use cases.
                        var useCases = descriptiveRow[titleCol].split(/\r?\n/);

                        for each (var useCase in useCases) {
                            description += "- " + useCase + "\n";
                        }

                        // Indicate that the use case has been imported.
                        var descriptiveStatusCell = sheet.getRange(descriptiveRowCount + 1, statusCol + 1, 1, 1);
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

                description += "\n**Hour Estimates**\n\n";


                // Fetch all roles with hours and append them to the description.
                for (var j = 0; j < roleCount; j++) {
                    if (hours[0][j] != "" && hours[0][j] > 0) {
                        description += "- **" + titles[0][j] + "** " + hours[0][j] + " hours\n";
                    }
                }

                description += "\n**Total Hour Estimate**\n\n- " + currentRow[hoursCol];

                // Set card title, description, point estimate, id, due date, assignees
                var card = createTrelloCard(currentRow[titleCol], description, currentRow[hoursCol], ScriptProperties.getProperty("listID"), dueDate, "");

                addTrelloLabels(card.id, currentRow[epicCol], existingLabels);

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

    var error = checkControlValues(false, true, false);
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
    var grid = app.createGrid(values.length + 1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i = values.length - 1; i >= 0; i--) {
        grid.setText(i + 1, 0, values[i].fullName);
        grid.setText(i + 1, 1, values[i].id);
    }

    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Members");

    SpreadsheetApp.getActiveSpreadsheet().show(app);

    return;
}

// Displays id's for boards which exist in your Trello account.
function displayBoards() {

    var error = checkControlValues(false, false, false);
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
    var grid = app.createGrid(values.length + 1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i = values.length - 1; i >= 0; i--) {
        grid.setText(i + 1, 0, values[i].name);
        grid.setText(i + 1, 1, values[i].id);
    }
    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Boards");

    SpreadsheetApp.getActiveSpreadsheet().show(app);

    return;
}

// Displays id's for checklists which exist in your Trello board.
function displayLists() {

    var error = checkControlValues(false, true, false);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var url = constructTrelloURL("boards/" + ScriptProperties.getProperty("boardID") + "/lists");
    var resp = UrlFetchApp.fetch(url, {"method": "get"});
    var values = Utilities.jsonParse(resp.getContentText())

    var app = UiApp.createApplication();

    var header1 = app.createHTML("<b>List Name</b>");
    var header2 = app.createHTML("<b>List Id</b>");
    var grid = app.createGrid(values.length + 1, 2).setWidth("100%");
    grid.setBorderWidth(5);
    grid.setWidget(0, 0, header1).setWidget(0, 1, header2);
    grid.setCellPadding(5);

    for (var i = values.length - 1; i >= 0; i--) {
        grid.setText(i + 1, 0, values[i].name);
        grid.setText(i + 1, 1, values[i].id);
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
