// Set up the Google Spreadsheets dropdown menu
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Display Available Members", functionName: "displayMembers"},{name: "Display Available Boards", functionName: "displayBoards"},{name: "Display Available Lists", functionName: "displayLists"},{name: "Upload Outstanding Backlog Items", functionName: "upload"}];
  ss.addMenu("Trello", menuEntries);
 }
 
function upload() {
    var startTime = new Date();  
    Logger.log("Started at:"+ startTime); 
    var error = checkControlValues(true,true);
    if (error != "") {
      Browser.msgBox("ERROR:Values in the Control sheet have not been set. Please fix the following error:\n " + error);
      return;
    }
    
    var url = constructTrelloURL("boards/"+ ScriptProperties.getProperty("boardID") + "/lists");
    var resp = UrlFetchApp.fetch(url, {"method": "get"});
    var lists = Utilities.jsonParse(resp.getContentText());
    var listIds = new Array();
    var listNames = new Array();
  
    for (var i=0; i< lists.length; i++) {
      listIds.push(lists[i].id);
      listNames.push(lists[i].name);
    } 
    
    
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backlog");
    var defaultListID = ScriptProperties.getProperty("listID");
    var boardID = ScriptProperties.getProperty("boardID");;
    var existingLabels = getExistingLabels(boardID);
    if (existingLabels == null || existingLabels.length ==0) {
      
      return;
    }
    var successCount = 0;
    var partialCount = 0;
  
    var rows=sheet.getDataRange().getValues();
               
    var headerRow = rows[0]; 
  
    for (var j = 1; j < rows.length; j++) {
      
      r=j+1;         
      var currentRow = rows[j];
      var status = currentRow[0];
      
      currentTime = new Date(); 
      Logger.log("Row " + r + ":" + currentTime); 
      if (currentRow[2].trim() == "") {
        // Do nothing if no card name
      }  
      else if (currentTime.valueOf() - startTime.valueOf() >= 330000) { // 5.5 minutes - scripts time out at 6 minutes
        Browser.msgBox("WARNING: Script was about to time out so upload has been terminated gracefully ." + successCount + " backlog items were uploaded successfully.");
        return;
      }
      else if (status == "Started") {
        Browser.msgBox("Error: Backlog item at row " + r + " has a status of 'Started' which means the Trello card MAY have been partially created for this item. Verify the state of the card, and either:\na) Delete the card from Trello if it's incomplete, and change status cell to blank.\n b)If card is complete, then change the status of the backlog item to 'Completed'");
        return;
      }  
      else if (status == "") {
        var listId = defaultListID;
        var overrideListName = currentRow[1];
        if (overrideListName != "") {
          var index = listNames.indexOf(overrideListName);
          if (index >= 0) {
            listId = listIds[index];
          }  
        }  
        if (listId == "") {
          Browser.msgBox("Could not determine list for row " + r + ". Aborting run.");
          return;
        }  
        var statusCell = sheet.getRange(r,1,1,1);
        var dueDate = null;
        if (currentRow[5] !== '') {
          dueDate = currentRow[5];
        }  
        statusCell.setValue("Started");
        partialCount ++;
        
        var card = createTrelloCard(currentRow[2],currentRow[3],currentRow[4],listId,dueDate,currentRow[7]);
        createTrelloAttachment(card.id,currentRow[8]);
        addTrelloLabels(card.id,currentRow[6],existingLabels);
        var comment = currentRow[9];
        var comments = comment.split("\n");
        
        for (var i = 0; i < comments.length; i++) {
          if (comments[i] != "") {
            createTrelloComment(card.id,comments[i]);
          }
        }
       
        for (var i = 11; i < headerRow.length; i++) {
          if (headerRow[i] !== "" && currentRow[i] !== "") {
            addChecklist(card, boardID,headerRow[i],currentRow[i]);
          }  
        }  
        
        statusCell.setValue("Completed");   
        SpreadsheetApp.flush();
        partialCount --;
        successCount ++;
          
      }
      else if (status != "Completed") {
          Browser.msgBox("Error: Backlog item at row " + r + " has a status of '" + status + "' Change status to 'Completed' if not required, or clear it to allow it to be uploaded." );
        return;
      }    
     
    }
    Browser.msgBox( successCount + " backlog items were uploaded successfully.");
    return;
}

function getExistingLabels(boardId) {

    var values = null;
    var url = constructTrelloURL("boards/" + boardId + "/labels");
    var resp = UrlFetchApp.fetch(url, {"method": "get","muteHttpExceptions":true });
    if (resp.getResponseCode() == 200) {
      var values = Utilities.jsonParse(resp.getContentText());
    }  
    else {
      Browser.msgBox("ERROR:Unable to return existing labels from board:" + resp.getContentText());
    }
      
    return values;
}  
 
 
function addChecklist(card, boardID,checklistName, checklistData) {
  
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
  

 
  
  
function createTrelloCard(cardName, cardDesc, storyPoints, listID, dueDate,members){
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
 
function createTrelloChecklist(cardID, name){
  var url = constructTrelloURL("checklists") + "&name=" + encodeURIComponent(name) + "&idCard="  + encodeURIComponent(cardID);
  var resp = UrlFetchApp.fetch(url, {"method": "post"});
  return Utilities.jsonParse(resp.getContentText());
}
  
function createTrelloComment(cardID, name){
  var url = constructTrelloURL("cards/"+ cardID + "/actions/comments") + "&text=" + encodeURIComponent(name);
  var resp = UrlFetchApp.fetch(url, {"method": "post"});
  return Utilities.jsonParse(resp.getContentText());
}

function createTrelloAttachment(cardID, attachment){
  if (attachment == "") {
    return;
  }  
  var attachments = attachment.split(",");
  for (var i= 0; i< attachments.length;i++) {
    var url = constructTrelloURL("cards/"+ cardID + "/attachments") + "&url=" + encodeURIComponent(attachments[i]);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});
  }  
  return;
}

function addTrelloMember(cardID, member){
  if (member == "") {
    return;
  }  
  var members = member.split(",");
  for (var i= 0; i< members.length;i++) {
    var url = constructTrelloURL("cards/"+ cardID + "/idMembers") + "&value=" + encodeURIComponent(members[i]);
    var resp = UrlFetchApp.fetch(url, {"method": "post"});
  }  
  return;
}
  
function addTrelloLabels(cardID, label,existingLabels){
  if (label == "" ) {
    return;
  }  
  var labels = label.split(",");
  for (var i= 0; i< labels.length;i++) {
    var labelId = getIdForLabelName(labels[i],existingLabels);
    if (labelId == null) {
      var url = constructTrelloURL("cards/"+ cardID + "/labels") + "&color=null&name=" + encodeURIComponent(labels[i]);
      var resp = UrlFetchApp.fetch(url, {"method": "post"});
    }
    else {
      var url = constructTrelloURL("cards/"+ cardID + "/idLabels") + "&value=" + encodeURIComponent(labelId);
      var resp = UrlFetchApp.fetch(url, {"method": "post"});

    }  
  }  
  return;
}

function getIdForLabelName(label,existingLabels) {
  
  for (var i=0; i < existingLabels.length;i++) {
    if (existingLabels[i].name.toUpperCase() == label.toUpperCase()) {
      return existingLabels[i].id;
    }  
  }  
  return null;
}  

function createTrelloChecklistItem(checkListID, name){
  var url = constructTrelloURL("checklists/" + checkListID + "/checkItems") + "&name=" + encodeURIComponent(name);
  var resp = UrlFetchApp.fetch(url, {"method": "post"});
  return Utilities.jsonParse(resp.getContentText());
}

function addTrelloChecklistToCard(checkListID, cardID) {
  var url = constructTrelloURL("cards/" + cardID + "/checklists") + "&value=" + encodeURIComponent(checkListID);
  var resp = UrlFetchApp.fetch(url, {"method": "post"});
  return Utilities.jsonParse(resp.getContentText());
}

 


  
function checkControlValues(requireList, requireBoard) {
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control").getRange("B3:B6").getValues();
  
  var appKey = col[0][0].toString().trim();
  if(appKey == "") {
    return "App Key not found";
  }  
  ScriptProperties.setProperty("appKey", appKey);
  
  var token = col[1][0].toString().trim();
  if(token == "") {
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
  
  return "";
  
} 
 
  
  
  
  



function constructTrelloURL(baseURL){
 
  return "https://trello.com/1/"+ baseURL +"?key="+ScriptProperties.getProperty("appKey")+"&token="+ScriptProperties.getProperty("token");
}

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
  
    
    for (var i=values.length-1;i>=0;i--) {
      grid.setText(i+1, 0, values[i].name);
      grid.setText(i+1, 1, values[i].id);
    }
    
    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Lists");
   
                     
    SpreadsheetApp.getActiveSpreadsheet().show(app);
  
  
  return;
}

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
  
    
    for (var i=values.length-1;i>=0;i--) {
      grid.setText(i+1, 0, values[i].name);
      grid.setText(i+1, 1, values[i].id);
    }
    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Boards");
   
                     
    SpreadsheetApp.getActiveSpreadsheet().show(app);
  
  
  return;
}

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
  
    
    for (var i=values.length-1;i>=0;i--) {
      grid.setText(i+1, 0, values[i].fullName);
      grid.setText(i+1, 1, values[i].id);
    }
    
    var panel = app.createScrollPanel(grid).setAlwaysShowScrollBars(true).setSize("100%", "100%");
    app.add(panel);
    app.setTitle("Available Members");
   
                     
    SpreadsheetApp.getActiveSpreadsheet().show(app);
  
  
  return;
}
