function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Tournaments")
    .addItem("Start", "setupInitial")
    .addItem("Add Names to Form", "addNamesToResults")
    .addSeparator()
    .addItem("New Stage", "newStage")
    .addItem("Make Sheets", "makeSheets")
    .addItem("Delete Stage", "deleteStage")
    .addSeparator()
    .addItem("Remove Result", "removeResult")
    .addItem("Test Webhook", "testWebhook")
    .addToUi();
}

function setupInitial() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet.getSheetByName('Administration')) {
    ui.alert("The setup for this tournament has already been started.", ui.ButtonSet.OK);
    return;
  }
  response = ui.alert("Would you like a signup form for your tournament?", ui.ButtonSet.YES_NO);
  var signupForm = false;
  if (response == ui.Button.YES) {
    signupForm = createSignups();
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var form = FormApp.create(sheet.getName() + ' Results Form');
  form.setDescription('Report results here - if you have previously played a partial match, please submit only new games').setConfirmationMessage("Thanks for participating!").setShowLinkToRespondAgain(false);
  var validation = FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build();
  form.addListItem().setTitle('Your username').setChoiceValues([""]).setRequired(true);
  form.addTextItem().setTitle('Your wins').setValidation(validation).setRequired(true);
  form.addListItem().setTitle('Opponent username').setChoiceValues([""]).setRequired(true);
  form.addTextItem().setTitle('Opponent wins').setValidation(validation).setRequired(true);
  form.addTextItem().setTitle('Comments');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();
  var formURL = 'https://docs.google.com/forms/d/' + form.getId() + '/viewform';

  var admin = sheet.insertSheet("Administration");
  admin.deleteColumns(5, 22);
  admin.deleteRows(7, 994);
  admin.setColumnWidth(1, 150);
  const prefill = "https://docs.google.com/forms/d/e/1FAIpQLSfZGOHcw0YhIsWD_dROOlvNdI8FBN9wuC-s7bitd31Ie1Xh1w/viewform?entry.1566745465=";
  let info = [['Tournament Name', sheet.getName(), '', ''],
              ['Published Spreadsheet', '', '=hyperlink("' + prefill + '"&B2, "Shortened:")', ''],
              ['Signup Form', signupForm ? signupForm : 'none', signupForm ? '=hyperlink("' + prefill + '"&B3, "Shortened:")' : '', ''],
              ['Results Form', formURL, '=hyperlink("' + prefill + '"&B4, "Shortened:")', ''],
              ['Discord Webhook', '', '=hyperlink("https://www.google.com/search?q=color+picker", "Color Hex:")', ''],
              ['','','',''],
              ['Stage', 'Type', 'Started','Finished']];
  admin.getRange(1, 1, 7, 4).setValues(info);
  admin.getRange(2, 3, 4, 1).setHorizontalAlignment('right');

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i=0; i < sheets.length; i++) {
    if (/^Form Responses/.test(sheets[i].getName())) {
      sheets[i].setName('Results');
      sheets[i].getRange(1, 7).setValue('Stage');
    }
  }

  var options = sheet.insertSheet("Options");
  options.deleteRows(2,999);
  options.deleteColumns(3, 24);
  options.setColumnWidth(1, 150);

  sheet.setActiveSheet(admin);
  ui.alert("Be sure to remember to add names to your results form (Tournaments > Add Names to Form) before starting the tournament.");
}

function addNamesToResults() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("What is the name of the sheet containing the list of players?");

  while (!sheet.getSheetByName(response.getResponseText()) && (response.getSelectedButton() == ui.Button.OK)) {
    response = ui.prompt("That sheet does not exist.", "What is the name of the sheet containing the list of players?", ui.ButtonSet.OK_CANCEL);
  }
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }

  var nameSheet = sheet.getSheetByName(response.getResponseText());
  var names = nameSheet.getRange(2, 2, nameSheet.getLastRow()-1, 1).getValues().flat();
  names.sort((a,b) => lowerSort(a,b));

  var resultsForm = FormApp.openByUrl(sheet.getSheetByName("Administration").getRange(4,2).getValue().replace(/viewform$/, "edit"));
  var questions = resultsForm.getItems();
  questions[0].asListItem().setChoiceValues(names);
  questions[2].asListItem().setChoiceValues(names);

  ui.alert("Names added to the results form. If you need to add more, you may either edit the form directly or rerun this function by going to Tournaments > Add Names to Form.")
}

function newStage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName("Administration");
  var stageList = [];
  if (admin.getLastRow() >= 8) {
    stageList = admin.getRange(8,1,admin.getLastRow()-7,4).getValues().map(x => x[0].toLowerCase());
  }
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("What is the name of the new stage?", ui.ButtonSet.OK_CANCEL);
  while (stageList.includes(response.getResponseText().toLowerCase()) && (response.getSelectedButton() == ui.Button.OK)) {
    response = ui.prompt("There is already a stage with that name.", "What is the name of the new stage?", ui.ButtonSet.OK_CANCEL);
  }
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var stageName = response.getResponseText();
  const stageTypes = ["single", "double", "swiss", "random", "group", "barrage"];
  response = ui.prompt("And what type of stage is " + stageName + "?", "Currently supported are: " + stageTypes.join(", "), ui.ButtonSet.OK_CANCEL);
  while (!stageTypes.includes(response.getResponseText().toLowerCase()) && (response.getSelectedButton() == ui.Button.OK)) {
    response = ui.prompt("That was not a valid stage type.", "What type of stage is " + stageName + "? Currently supported are: " + stageTypes.join(", "), ui.ButtonSet.OK_CANCEL);
  }
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var stageType = response.getResponseText().toLowerCase();

  admin.appendRow([stageName, stageType]);
  var stageNumber = admin.getLastRow() - 8;
  admin.getRange(stageNumber + 8, 3, 1, 2).insertCheckboxes()

  var options = sheet.getSheetByName("Options");
  options.insertRows(options.getLastRow()+1, 10);
  switch (stageType) {
      case 'single':
        singleBracketOptions(stageNumber, stageName);
        break;
      case 'double':
        doubleBracketOptions(stageNumber, stageName);
        break;
      case 'swiss':
        swissOptions(stageNumber, stageName);
        break;
      case 'random':
        randomOptions(stageNumber, stageName);
        break;
      case 'group':
        groupsOptions(stageNumber, stageName);
        break;
      case 'barrage':
        barrageOptions(stageNumber, stageName);
        break;
    }

    var toprow = options.getRange(10*stageNumber+1,1, 1, 2)
    toprow.setBackground('#D0E0E3').merge();
    sheet.setActiveSheet(options);
    options.setActiveRange(toprow);
    ui.alert("Fill out the options for the stage, then go to Tournaments > Make Sheets to set up the sheets for it.");
}

function makeSheets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName("Administration");
  var stageList = admin.getRange(8,1,admin.getLastRow()-7,4).getValues();
  var stageNames = stageList.map(x => x[0]).concat([""]);

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Which stage would you like to (re)make the sheets for?", "You may leave this blank to do so for all stages which have not started.", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  while (!stageNames.includes(response.getResponseText()) && (response.getSelectedButton() == ui.Button.OK)) {
    response = ui.prompt("That stage does not exist.", "Which stage would you like to (re)make the sheets for? You may leave this blank to do so for all stages which have not started.", ui.ButtonSet.OK_CANCEL);
  }
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  if (response.getResponseText() === "") {
    var remade = [];
    for (let i=0; i<stageList.length; i++) {
      if (!stageList[i][2]) {
        removeStage(stageList[i][0], stageList[i][1]);
        addStage(i, stageList[i][0], stageList[i][1]);
        remade.push(stageList[i][0]);
      }
    }
  } else {
    var stageName = response.getResponseText();
    var stageNumber = stageNames.indexOf(response.getResponseText());
    var stageType = stageList[stageNumber][1];
    if (stageList[stageNumber][2]) {
      ui.alert("Cannot make sheets: the stage you inputted has already started.");
      return;
    }
    removeStage(stageName, stageType);
    addStage(stageNumber, stageName, stageType);
  }
  
  ui.alert("Sheets made for " + ((response.getResponseText() === "") ? "stages: " + remade.join(", ") : "the " + stageName + " stage") + ".", "If you change your mind about an option, simply change it in the Options sheet and rerun this function by going to Tournaments > Make Sheets.", ui.ButtonSet.OK)
}

function deleteStage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName("Administration");
  var options = sheet.getSheetByName("Options");
  var stageList = admin.getRange(8,1,admin.getLastRow()-7,4).getValues();
  var stageNames = stageList.map(x => x[0]);

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Which stage would you like to delete?", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  while (!stageNames.includes(response.getResponseText()) && (response.getSelectedButton() == ui.Button.OK)) {
    response = ui.prompt("That stage does not exist.", "Which stage would you like to delete?", ui.ButtonSet.OK_CANCEL);
  }
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var stageName = response.getResponseText();
  var stageNumber = stageNames.indexOf(response.getResponseText());
  var stageType = stageList[stageNumber][1];
  if (stageList[stageNumber][2]) {
    ui.alert("Cannot delete stage: the stage you inputted has already started.");
    return;
  }
  admin.deleteRow(stageNumber + 8);
  options.deleteRows(10*stageNumber+1, 10);
  removeStage(stageName, stageType);
  ui.alert(stageName + " has been deleted.")
}

function addStage(num, name, type) {
  switch (type) {
      case 'single':
        createSingleBracket(num, name);
        break;
      case 'double':
        createDoubleBracket(num, name);
        break;
      case 'swiss':
        createSwiss(num, name);
        break;
      case 'random':
        createRandom(num, name);
        break;
      case 'group':
        createGroups(num, name);
        break;
      case 'barrage':
        createBarrage(num, name);
        break;
    }
}

function removeStage(name, type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  switch (type) {
        case 'single':
          if (sheet.getSheetByName(name + ' Bracket')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Bracket'));}
          if (sheet.getSheetByName(name + ' Places')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Places'));}
          break;
        case 'double':
          if (sheet.getSheetByName(name + ' Bracket')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Bracket'));}
          if (sheet.getSheetByName(name + ' Places')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Places'));}
          break;
        case 'swiss':
          if (sheet.getSheetByName(name + ' Standings')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Standings'));}
          if (sheet.getSheetByName(name + ' Processing')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Processing'));}
          break;
        case 'random':
          if (sheet.getSheetByName(name + ' Standings')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Standings'));}
          if (sheet.getSheetByName(name + ' Processing')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Processing'));}
          break;
        case 'group':
          if (sheet.getSheetByName(name + ' Standings')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Standings'));}
          if (sheet.getSheetByName(name + ' Processing')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Processing'));}
          var sheets = sheet.getSheets();
          let aggRegex = new RegExp('^' + name + ' Aggregation');
          for (let j=sheets.length-1; j>=0; j--) {
            if(aggRegex.test(sheets[j].getSheetName())) {
              sheet.deleteSheet(sheets[j]);
              sheets.splice(j,1);
            }
          }
          break;
        case 'barrage':
          if (sheet.getSheetByName(name + ' Matches')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Matches'));}
          if (sheet.getSheetByName(name + ' Places')) {sheet.deleteSheet(sheet.getSheetByName(name + ' Places'));}
          break;
      }
}

function removeResult() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Enter the row number of the result you wish to remove.", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var rowToRemove = response.getResponseText();

  if (/^\d+$/.test(rowToRemove)) {
    rowToRemove = Number(rowToRemove);
  } else {
    ui.alert("Non-number entered");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues()
  var stageList = admin.slice(7);
  var stageNames = stageList.map(x => x[0]);
  var removal = sheet.getSheetByName('Results').getRange(rowToRemove,1,1,8).getValues()[0];
  var success = false;

  var stageName = removal[6];
  var stageNumber = stageNames.indexOf(stageName);

  if (stageNumber >= 0) {
    if (stageList[stageNumber][2] && !stageList[stageNumber][3]) {
      switch (stageList[stageNumber][1]) {
        case 'single':
          success = removeSingleBracket(stageNumber, stageName, rowToRemove);
          break;
        case 'double':
          success = removeDoubleBracket(stageNumber, stageName, rowToRemove);
          break;
        case 'swiss':
          success = removeSwiss(stageNumber, stageName, rowToRemove);
          break;
        case 'random':
          success = removeRandom(stageNumber, stageName, rowToRemove);
          break;
        case 'group':
          success = removeGroups(stageNumber, stageName, rowToRemove);
          break;
        case 'barrage':
          success = removeBarrage(stageNumber, stageName, rowToRemove);
          break;
      }
      if (success) {
        var form = FormApp.openByUrl(admin[3][1].replace(/viewform$/, 'edit'));
        var items = form.getItems();
        var zeroMatch = form.createResponse();
        zeroMatch.withItemResponse(items[0].asListItem().createResponse(removal[1])).withItemResponse(items[1].asTextItem().createResponse('0'));
        zeroMatch.withItemResponse(items[2].asListItem().createResponse(removal[3])).withItemResponse(items[3].asTextItem().createResponse('0'));
        zeroMatch.withItemResponse(items[4].asTextItem().createResponse('Removal ' + rowToRemove + ' - ' + removal[6] + ': ' + removal[1] + ' ' + removal[2] + '-' + removal[4] + ' ' + removal[3]));
        zeroMatch = zeroMatch.toPrefilledUrl().replace("viewform", "formResponse") + "&submit=Submit";
        UrlFetchApp.fetch(zeroMatch);
      }
    } else {
      ui.alert('Removal Failure: the result is from an inactive stage.');
    }
  } else {
    ui.alert('Removal Failure: the result is not in any stage.');
  }
}

function testWebhook() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[4][1];

  if (webhook) {    
    var content = "This webhook will display results as they come in. Here are some useful links for this tournament:";
    content += "\n\n**[Match Information and Results](" + admin[1][1] + ")**"
    if (admin[2][1] != 'none') {
      content += "\n**[Signup Form](" + admin[2][1] + ")**";
    }
    content += "\n**[Results Form](" + admin[3][1] + ")**\n\nBest of luck to all!"
    
    var testMessage = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": null,
        "embeds": [
          {
            "title": "Welcome to " + admin[0][1],
            "color": parseInt('0x'.concat(admin[4][3].slice(1))),
            "description": content
          }
        ]
      })
    };
    
    UrlFetchApp.fetch(webhook, testMessage);
  } else {
    SpreadsheetApp.getUi().alert("No webhook URL provided");
  }
  
}

function createSignups() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create(sheet.getName() + ' Signup Form');
  form.setDescription("Sign up for " + sheet.getName() + " here.").setConfirmationMessage("Thanks for signing up!").setShowLinkToRespondAgain(false);
  form.addTextItem().setTitle("Dominion Online Username").setRequired(true);
  form.addTextItem().setTitle("Discord Username (if different)");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  ScriptApp.newTrigger('onSignup').forForm(form).onFormSubmit().create();
  var formURL = 'https://docs.google.com/forms/d/' + form.getId() + '/viewform';

  sheet.renameActiveSheet(sheet.getActiveSheet().getName());
  var sheets = sheet.getSheets();
  for (let i=0; i < sheets.length; i++) {
    if (/^Form Responses/.test(sheets[i].getName())) {
      sheets[i].setName('Signups');
    }
  }

  return formURL;
}

function onFormSubmit() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration');
  var stageList = admin.getRange(8,1,admin.getLastRow()-7,4).getValues();
  var found = false;
  
  for (let i=stageList.length-1; i>=0; i--) {
    if ((stageList[i][2]) && (!stageList[i][3])) {
      switch (stageList[i][1]) {
        case 'single':
          found = updateSingleBracket(i, stageList[i][0]);
          break;
        case 'double':
          found = updateDoubleBracket(i, stageList[i][0]);
          break;
        case 'swiss':
          found = updateSwiss(i, stageList[i][0]);
          break;
        case 'random':
          found = updateRandom(i, stageList[i][0]);
          break;
        case 'group':
          found = updateGroups(i, stageList[i][0]);
          break;
        case 'barrage':
          found = updateBarrage(i, stageList[i][0]);
          break;
      }
    }
    if (found) {break;}
  }
  
  if (!found) {
    sendResultToDiscord('', 'unfound', 0);
    Utilities.sleep(2000);
    var results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();    
    results.getRange(resultsLength, 8).setValue('not found');
  }
}

function sendResultToDiscord(name, displayType, gid) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[4][1];

  var form = FormApp.openByUrl(admin[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  
  var display = '**' + mostRecent[0] + ' ' + mostRecent[1] + '-' + mostRecent[3] + ' ' + mostRecent[2] + '**';
  if (displayType == 'unfound') {
    display += '\nThis results submission did not match any active matches. If it should have, the tournament organizer should make sure that active stages are properly checked as having started and that there are no spreadsheet errors.';
    if (webhook) { 
      var unfoundEmbed = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
        },
        "payload": JSON.stringify({
          "content": null,
          "embeds": [{
            "title": "Invalid Result Submitted",
            "color": parseInt('0x'.concat(admin[4][3].slice(1))),
            "description": display
          }]
        })
      };
      UrlFetchApp.fetch(webhook, unfoundEmbed);
    } else {
      Logger.log(display.replace('\n', ': '));
    }
  } else if (webhook) {
    if (mostRecent[4]) {
      display += '\n' + mostRecent[4];
    }
    display += '\n[Current ' + name + ' ' + displayType + '](' + admin[1][1] + '?gid=' + gid + ')';
    
    var scoreEmbed = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": null,
        "embeds": [{
          "title": "New " + admin[0][1] + " " + name + " Result",
          "color": parseInt('0x'.concat(admin[4][3].slice(1))),
          "description": display
        }]
      })
    };
    var matchCommand = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": "!match " + mostRecent[0] + ", " + mostRecent[2] + " -c"
      })
    };
    
    UrlFetchApp.fetch(webhook, scoreEmbed);
    UrlFetchApp.fetch(webhook, matchCommand);
  }  
}

function onSignup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var signups = sheet.getSheetByName("Signups")
  var webhook = admin[4][1];

  var form = FormApp.openByUrl(admin[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});

  if (webhook) {
    var message = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": null,
        "embeds": [
          {
            "title": mostRecent[0] + " has signed up!",
            "color": parseInt('0x'.concat(admin[4][3].slice(1))),
            "description": "That makes " + (signups.getLastRow() - 1) + " players who want to play in " + sheet.getName() + "\n\n**[Full List of Signups](" + admin[1][1] + "?gid=" + signups.getSheetId() + ")**"
          }
        ]
      })
    };

    UrlFetchApp.fetch(webhook, message);
  }
}

function shuffle(arr, start=0, end=arr.length-1) {
  for(let i=end-start; i>0; i--) {
    let toend = Math.floor(Math.random()*(i+1) + start);
    [arr[i+start],arr[toend]] = [arr[toend], arr[i+start]];
  }
}

function columnFinder(n) {
  var letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  if (n <= 26) {
    return letters[n-1];
  } else {
    return letters[Math.floor((n-1)/26)-1] + letters[((n-1)%26)];
  }
}

function lowerSort(a,b) {
  if (a.toLowerCase() < b.toLowerCase()) {
    return -1;
  } else if (a.toLowerCase() > b.toLowerCase()) {
    return 1;
  } else if (a.toLowerCase() == b.toLowerCase()) {
    return 0
  }  
}

// Single

function singleBracketOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,9,2).setValues([
    ["''" + name + "' Single Elimination Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Games to Win", ""],
    ["Seeding Method", "standard"],
    ["Reseeding", ""],
    ["Consolation", "none"],
    ["Places Sheet", ""],
    ["Bracket Count", "1"]
  ]);

  options.getRange(base + 5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'tennis'],true).build());
  options.getRange(base + 6,2).insertCheckboxes();
  options.getRange(base + 7,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['none', 'third'],true).build());
  options.getRange(base + 8,2).insertCheckboxes();
}

function createSingleBracket(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[8][1]);
  var nround = Math.ceil(Math.log(nplayers/nbrackets)/Math.log(2));
  var bracketSize = 2**(nround + 1);
  var bracket = sheet.insertSheet(name + ' Bracket');
  bracket.deleteRows(2, 999);
  if ((nbrackets == 1) && (options[6][1] == 'third') && (nround < 4)) {
    bracket.insertRows(1, 3*bracketSize/4 + 5)
  } else {
    bracket.insertRows(1, nbrackets*bracketSize - 1);
  }
  bracket.deleteColumns(2, 25);
  bracket.insertColumns(1, 2*nround + 1);
  bracket.getRange(bracket.getMaxRows(), bracket.getMaxColumns()).setFontColor("#FFFFFF").setValue("~");
  
  //make borders
  for (let b=0; b < nbrackets; b++) {
    for (let j=1; j <= nround; j++) {
      let nmatches = 2**(nround-j);
      let space = 2**j;
      for (let i=0; i < nmatches; i++) {
        bracket.getRange(b*bracketSize + space + 2*i*space, 2*j, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        if (!options[5][1] && (j != 1)) {
          bracket.getRange(b*bracketSize + space/2 + 2*i*space, 2*j, space, 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
      }
    }
  }
  
  if ((nbrackets == 1) && (options[6][1] == 'third')) {
    let top = 2**nround + 2**(nround-1) + 4
    bracket.getRange(top, 2*nround, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    bracket.getRange(top-1, 2*nround).setFontColor('#b7b7b7').setValue('Third Place Match');
  }
  
  //column widths
  bracket.setColumnWidth(1, 50)
  for (let j=2; j <= 2*nround + 1; j++) {
    if (j % 2) {
      bracket.setColumnWidth(j, 30);
    } else {
      bracket.setColumnWidth(j, 150);
    }
  }
  
  //seed numbers and players
  if (options[5][1]) {
    let storeRow = [];
    for (let j=1; j<=nround-1; j++) {
      storeRow.push("");
      storeRow.push(nbrackets*(2**(nround-j)));
    }
    storeRow[1] -= (nbrackets*(2**nround) - nplayers);
    bracket.getRange(1,2,1,storeRow.length).setFontColor("#FFFFFF").setValues([storeRow]);
    reseed(num, name, 1);
  } else {
    var allSeeds = Array.from({length: nplayers}, (v, i) => (i+1));
    var seedMethod = options[4][1];
    switch (seedMethod) {
      case 'random':
        shuffle(allSeeds);
        allSeeds = allSeeds.concat(Array.from({length: nbrackets * 2**nround - nplayers}, () => 0));
        break;
      case 'tennis':
        let lastWithBye = nbrackets * 2**nround - nplayers;
        allSeeds = allSeeds.concat(Array.from({length: nbrackets * 2**nround - nplayers}, () => 0));
        shuffle(allSeeds, nbrackets * 2**(nround-1));      
        for (let i=nround-1; i>0; i--) {
          shuffle(allSeeds, nbrackets * 2**(i-1), nbrackets * 2**i - 1);
        }
        let zeros = [];
        let idx = allSeeds.indexOf(0);
        while(idx != -1) {
          zeros.push(idx);
          idx = allSeeds.indexOf(0, idx + 1);
        }
        for (let i=0; i<nplayers; i++) {
          Logger.log(zeros);
          if (!zeros.length) {
            break;
          }
          if (allSeeds[i] <= lastWithBye) {
            let included = zeros.indexOf(allSeeds.length - i - 1)
            if (included != -1) {
              zeros.splice(included,1);
            } else {
              let toswap = Math.floor(zeros.length*Math.random());
              let k = allSeeds[allSeeds.length - i - 1];
              allSeeds[allSeeds.length - i - 1] = allSeeds[zeros[toswap]];
              allSeeds[zeros[toswap]] = k;
              zeros.splice(toswap,1);
            }
          }
        }
        break;
      default:
        allSeeds = allSeeds.concat(Array.from({length: nbrackets * 2**nround - nplayers}, () => 0));
        break;
    }

    var seedList = Array.from({length: nbrackets}, () => []);
    let nseeded = allSeeds.length
    for (let i=0; i<nseeded; i++) {
      if (Math.floor(i/nbrackets) % 2) {
        seedList[nbrackets - 1 - (i%nbrackets)].push(allSeeds[i]);
      } else {
        seedList[i%nbrackets].push(allSeeds[i]);
      }
    }
    
    function placeSeeds(seed, b, location, roundsLeft) {    
      if (seedList[b][2**nround - seed]) {
        if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location, 1).setValue(seedList[b][seed-1]);}
        bracket.getRange(b*bracketSize + location, 2).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[b][seed-1] + 1) + ")");
        if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location + 1, 1).setValue(seedList[b][2**nround - seed]);}
        bracket.getRange(b*bracketSize + location + 1, 2).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[b][2**nround - seed] + 1) + ")");
      } else {
        bracket.getRange(b*bracketSize + location, 2, 2, 2).setBorder(false, false, false, false, null, null);
        if ((location - 2) % 8) {
          if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location - 1, 3).setValue(seedList[b][seed-1]);}
          bracket.getRange(b*bracketSize + location - 1, 4).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[b][seed-1] + 1) + ")");        
        } else {
          if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location + 2, 3).setValue(seedList[b][seed-1]);}
          bracket.getRange(b*bracketSize + location + 2, 4).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[b][seed-1] + 1) + ")");        
        }      
      }
      
      for(let i=nround-roundsLeft+1; i < nround; i++) {
        placeSeeds(2**i + 1 - seed, b, location + 2**(nround+1-i) , nround - i);
      }
      
    }
    for (let b=0; b<nbrackets; b++) {placeSeeds(1,b,2,nround);}
  }
  
  
  //places sheet if desired
  if (options[7][1]) {
    var places = sheet.insertSheet(name + ' Places');
    places.deleteColumns(3,24);
    places.setColumnWidth(1,40);
    places.setColumnWidth(2,150);
    places.getRange('A:A').setVerticalAlignment('top').setHorizontalAlignment('right');
    places.getRange(1,1,2,2).setValues([
      ['Place', 'Player'],
      [1, '']
    ]);
    places.getRange(2,1,nbrackets,1).merge();
    places.getRange(nbrackets+2,1).setValue(2);
    places.getRange(nbrackets+2,1,nbrackets,1).merge();
    for (let n=2; n<nplayers; n*=2) {
      places.getRange(n*nbrackets+2,1).setValue('T' + (n+1));
      places.getRange(n*nbrackets+2,1,n*nbrackets,1).merge();
    }
    if ((nbrackets == 1) && (options[6][1] == 'third')) {
      places.getRange(4,1,2,1).breakApart().setValues([[3],[4]]);
    }
    places.deleteRows(nplayers + 2, 999-Number(nplayers));
  }
}

function updateSingleBracket(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse();});
  var bracket = sheet.getSheetByName(name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[8][1]);
  var nround = Math.ceil(Math.log(nplayers/nbrackets)/Math.log(2));
  var bracketSize = 2**(nround + 1);
  var gamesToWin = options[3][1];
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
  var found = false;
  
  for (let b=0; b<nbrackets; b++) {
    for (let j=nround; j>0; j--) {
      let nmatches = 2**(nround-j);
      let space = 2**j;
      for (let i=0; i<nmatches; i++) {
        if ((bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).getValue() === '') || (j == nround) || options[5][1]) {
          let match = bracket.getRange(b*bracketSize + space + 2*i*space, 2*j, 2, 2).getValues();
          if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
            var topwins = Number(match[0][1]) + Number(mostRecent[1]);
            var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
            found = true;
          } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
            var topwins = Number(match[0][1]) + Number(mostRecent[3]);
            var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
            found = true;
          }
          if (found) {
            bracket.getRange(b*bracketSize + space + 2*i*space, 2*j+1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            var loc = b + ',' + j + ',' + i;
            if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
              if (j != nround) {
                if (options[5][1] && !((j == nround - 1) && (nbrackets == 1))) {
                  let winList = bracket.getRange(1, 2*j).getValue();
                  bracket.getRange(1, 2*j).setValue(winList + "~" + match[0][0] + "~");
                  reseed(num, name, j+1);
                } else {
                  bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[0][0]);
                }
                if ((j == nround - 1) && (nbrackets == 1) && (options[6][1] == 'third')) {
                  bracket.getRange(4 + 3*space + ((i==1)?1:0), 2*j+2).setValue(match[1][0]);
                } else if (options[7][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let fillLength = nbrackets*nmatches;
                  let tofill = places.getRange(fillLength+2,2,fillLength,1).getValues().flat();
                  for (let k=0; k<fillLength; k++) {
                    if (!tofill[k]) {
                      places.getRange(fillLength+2+k,2).setValue(match[1][0]);
                      break;
                    }
                  }
                }
              } else if (options[7][1]) {
                let places = sheet.getSheetByName(name + ' Places');
                let tofill = places.getRange(2,2,nbrackets,1).getValues().flat();
                for (let k=0; k<nbrackets; k++) {
                  if (!tofill[k]) {
                    places.getRange(2+k,2).setValue(match[0][0]);
                    break;
                  }
                }
                tofill = places.getRange(nbrackets+2,2,nbrackets,1).getValues().flat();
                for (let k=0; k<nbrackets; k++) {
                  if (!tofill[k]) {
                    places.getRange(nbrackets+2+k,2).setValue(match[1][0]);
                    break;
                  }
                }
              }
              bracket.getRange(b*bracketSize + space + 2*i*space, 2*j, 1, 2).setBackground('#D9EBD3');
            } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
              if (j != nround) {
                if (options[5][1] && !((j == nround - 1) && (nbrackets == 1))) {
                  let winList = bracket.getRange(1, 2*j).getValue();
                  bracket.getRange(1, 2*j).setValue(winList + "~" + match[1][0] + "~");
                  reseed(num, name, j+1);
                } else {
                  bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[1][0]);
                }
                if ((j == nround - 1) && (nbrackets == 1) && (options[6][1] == 'third')) {
                  bracket.getRange(4 + 3*space + ((i==1)?1:0), 2*j+2).setValue(match[0][0]);
                } else if (options[7][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let fillLength = nbrackets*nmatches;
                  let tofill = places.getRange(fillLength+2,2,fillLength,1).getValues().flat();
                  for (let k=0; k<fillLength; k++) {
                    if (!tofill[k]) {
                      places.getRange(fillLength+2+k,2).setValue(match[0][0]);
                      break;
                    }
                  }
                }
              } else if (options[7][1]) {
                let places = sheet.getSheetByName(name + ' Places');
                let tofill = places.getRange(2,2,nbrackets,1).getValues().flat();
                for (let k=0; k<nbrackets; k++) {
                  if (!tofill[k]) {
                    places.getRange(2+k,2).setValue(match[1][0]);
                    break;
                  }
                }
                tofill = places.getRange(nbrackets+2,2,nbrackets,1).getValues().flat();
                for (let k=0; k<nbrackets; k++) {
                  if (!tofill[k]) {
                    places.getRange(nbrackets+2+k,2).setValue(match[0][0]);
                    break;
                  }
                }
              }
              bracket.getRange(b*bracketSize + space + 2*i*space+1, 2*j, 1, 2).setBackground('#D9EBD3');
            }
            break;
          } else if ((j == nround) && (nbrackets == 1) && (options[6][1] == 'third')) {
            match = bracket.getRange(4 + 1.5*space, 2*j, 2, 2).getValues();
            if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[1]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
              found = true;
            } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[3]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
              found = true;
            }
            if (found) {
              bracket.getRange(4 + 1.5*space, 2*j+1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
              var loc = 'third';
              if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                bracket.getRange(4 + 1.5*space, 2*j, 1, 2).setBackground('#D9EBD3');
                if (options[7][1]) {sheet.getSheetByName(name + ' Places').getRange(4,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
              } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                bracket.getRange(5 + 1.5*space, 2*j, 1, 2).setBackground('#D9EBD3');
                if (options[7][1]) {sheet.getSheetByName(name + ' Places').getRange(4,2,2,1).setValues([[match[1][0]],[match[0][0]]]);}
              }
              break;
            }
          }
        }
      }
      if (found) {break;}
    }
    if (found) {break;}
  }
  
  if (found && !empty) {
    sendResultToDiscord(name, 'Bracket', bracket.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(loc);
  }
  
  return found;
}

function reseed(num, name, round) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var bracket = sheet.getSheetByName(name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[8][1]);
  var nround = Math.ceil(Math.log(nplayers/nbrackets)/Math.log(2));
  var bracketSize = 2**(nround + 1);
  var byes = nbrackets*(2**nround) - nplayers;
  let seedMethod = options[4][1];

  if (round == 1) {
    let nmatches = nbrackets*(2**(nround-1));
    let allSeeds = Array.from({length: nplayers}, (v, i) => (i+1));
    switch (seedMethod) {
      case 'random':
        shuffle(allSeeds);
        break;
      case 'tennis':
        shuffle(allSeeds, nmatches)
        break;
    }

    for (let i=0; i<byes; i++) {
      bracket.getRange(nbrackets*bracketSize-2-4*i, 2, 2, 2).setBorder(false, false, false, false, null, null);
      bracket.getRange(4+8*i, 3, 1, 2).setValues([[allSeeds[i], "=index('" + options[1][1] + "'!B:B," + (allSeeds[i]+1) + ")"]]);
    }
    Logger.log(allSeeds);
    for (let i=byes; i<nmatches; i++) {
      bracket.getRange(2+4*(i-byes), 1, 2, 2).setValues([[allSeeds[i], "=index('" + options[1][1] + "'!B:B," + (allSeeds[i]+1) + ")"],
      [allSeeds[2*nmatches-i-1], "=index('" + options[1][1] + "'!B:B," + (allSeeds[2*nmatches-i-1]+1) + ")"]]);
    }
  } else {
    let winnerInfo = bracket.getRange(1,2*round-2,1,2).getValues()[0];
    let winners = winnerInfo[0].slice(1,-1).split("~~");
    if (winners.length != Number(winnerInfo[1])) {
      return;
    }
    let playerList = sheet.getSheetByName(options[1][1]).getRange(2,2,nplayers,1).getValues().flat()
    if ((round == 2) && (byes > 0)) {
      winners = winners.concat(playerList.slice(0,byes));
    }
    let nmatches = nbrackets*(2**(nround-round));
    let seedList = winners.map(pl => [playerList.indexOf(pl) + 1, pl]);
    seedList.sort((a,b) => a[0] - b[0]);
    switch (seedMethod) {
      case 'random':
        shuffle(seedList);
        break;
      case 'tennis':
        shuffle(seedList, nmatches)
        break;
    }
    let space = 2**round;
    for (let i=0; i<nmatches; i++) {
      bracket.getRange(space + 2*i*space, 2*round-1, 2, 2).setValues([seedList[i], seedList[2*nmatches-i-1]])
    }
  }
}

function removeSingleBracket(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',');
  var bracket = sheet.getSheetByName(name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[8][1]);
  var nround = Math.ceil(Math.log(nplayers/nbrackets)/Math.log(2));
  var bracketSize = 2**(nround + 1);
  
  if (loc == 'third') {
    var match = bracket.getRange(2**nround + 2**(nround-1) + 4, 2*nround, 2, 2).setBackground(null).getValues();
    if (removal[1] == match[0][0]) {
      var topwins = Number(match[0][1]) - Number(removal[2]);
      var bottomwins = Number(match[1][1]) - Number(removal[4]);
    } else {
      var topwins = Number(match[0][1]) - Number(removal[4]);
      var bottomwins = Number(match[1][1]) - Number(removal[2]);
    }
    bracket.getRange(3 * 2**(nround-1) + 4, 2*nround+1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
    if (options[7][1]) {
      sheet.getSheetByName(name + ' Places').getRange(4,2,2,1).setValues([[''],['']]);
    }
  } else {
    var space = 2**loc[1];
    
    //scrub dependencies on this match and fail if one has a result
    if (options[5][1]) {
      if (bracket.getRange(2*space + 1, 2*loc[1]+2).getValue()) {
        let nmatches = nbrackets * 2**(nround - loc[1] - 1);
        for (let i=0; i<nmatches; i++) {
          if (bracket.getRange(2*space + 4*i*space, 2*loc[1]+3).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Reseeded Single Elimination): The next round has a result');
            return false;
          }
        }
        for (let i=0; i<nmatches; i++) {
          bracket.getRange(2*space + 4*i*space, 2*loc[1]+1, 2, 2).setValues([['', ''], ['', '']]);
        }
      }
      let winners = bracket.getRange(1,2*loc[1]).getValue();
      winners = winners.split("~" + removal[1] + "~").join('');
      winners = winners.split("~" + removal[3] + "~").join('');
      bracket.getRange(1,2*loc[1]).setValue(winners);
    } else {
      var toScrub = [];
      if (bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+3).getValue() !== '') {
        SpreadsheetApp.getUi().alert('Removal Failure (Single Elimination): The match the winner plays in has a result');
        return false;
      } else {
        toScrub.push(bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1] + 2));
      }
      if ((loc[1] == nround-1) && (nbrackets == 1) && (options[6][1] == 'third')) {
        if (bracket.getRange(4 + 3*space, 2*nround+1).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Single Elimination): The third place match has a result');
          return false;
        } else {
          toScrub.push(bracket.getRange(4 + 3*space + ((loc[2]==1) ? 1 : 0), 2*loc[1] + 2));
        }
      }
      for (let i=toScrub.length; i>0; i--) {toScrub[i-1].setValue('');}
    }
    
    var match = bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2], 2*loc[1], 2, 2).setBackground(null).getValues();
    if (removal[1] == match[0][0]) {
      var topwins = Number(match[0][1]) - Number(removal[2]);
      var bottomwins = Number(match[1][1]) - Number(removal[4]);
    } else {
      var topwins = Number(match[0][1]) - Number(removal[4]);
      var bottomwins = Number(match[1][1]) - Number(removal[2]);
    }
    bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2], 2*loc[1] + 1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
    if (options[7][1]) {
      let places = sheet.getSheetByName(name + ' Places');
      let nmatches = 2**(nround-loc[1]);
      let fillLength = nbrackets*nmatches;
      let tofill = places.getRange(2+fillLength,2,fillLength,1).getValues();
      for (let i=0; i<fillLength; i++) {
        if ((tofill[i] == match[0][0]) || (tofill[i] == match[1][0])) {
          places.getRange(2+fillLength+i,2).setValue('');
        }
      }
      if (loc[1] == nround) {
        tofill = places.getRange(2,2,fillLength,1).getValues();
        for (let i=0; i<fillLength; i++) {
          if ((tofill[i] == match[0][0]) || (tofill[i] == match[1][0])) {
            places.getRange(2+i,2).setValue('');
          }
        }
      }
    }
  }
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

// Double

function doubleBracketOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,7,2).setValues([
    ["''" + name + "' Double Elimination Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Games to Win", ""],
    ["Seeding Method", "standard"],
    ["Finals", "standard"],
    ["Places Sheet", ""]
  ]);

  options.getRange(base + 5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'tennis'],true).build());
  options.getRange(base + 6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'grand'],true).build());
  options.getRange(base + 7,2).insertCheckboxes();
}

function createDoubleBracket(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var nlround = 2*nround - 2 - (2**nround-options[2][1]>=2**(nround-2));
  var width = Math.max(2*nround + 5, 2*nlround + 1)
  var bracket = sheet.insertSheet(name + ' Bracket');
  bracket.deleteRows(2, 999);
  bracket.insertRows(1, 2**(nround+1) + ((nlround % 2) ? 4 * 2**((nlround-1)/2) : 3 * 2**(nlround/2)));
  bracket.deleteColumns(2, 25);
  bracket.insertColumns(1, width);
  bracket.getRange(bracket.getMaxRows(), bracket.getMaxColumns()).setFontColor("#FFFFFF").setValue("~");
  
  //winners' bracket
  //make borders and labels
  for (let j=1; j <= nround; j++) {
    let nmatches = 2**(nround-j);
    let space = 2**j;
    for (let i=0; i < nmatches; i++) {
      bracket.getRange(space + 2*i*space, 2*j, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      bracket.getRange(space + 2*i*space - 1, 2*j).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue('Match ' + j + '|' + (i+1));
      if (j != 1) {
        bracket.getRange(space/2 + 2*i*space, 2*j, space, 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }   
    }
  }
  
  //column widths
  bracket.setColumnWidth(1, 50);
  bracket.getRange('A:A').setHorizontalAlignment('right');
  for (let j=2; j <= width; j++) {
    if (j % 2) {
      bracket.setColumnWidth(j, 30);
    } else {
      bracket.setColumnWidth(j, 150);
    }
  }
  
  //seed numbers and players
  var seedList = Array.from({length: options[2][1]}, (v, i) => (i+1));
  var seedMethod = options[4][1];
  switch (seedMethod) {
    case 'random':
      shuffle(seedList);
      break;
    case 'tennis':
      let lastWithBye = 2**nround - options[2][1];
      let splitRound = lastWithBye ? Math.ceil(Math.log(lastWithBye)/Math.log(2)) : -1;
      shuffle(seedList, 2**(nround-1));      
      for (let i=nround-1; i>1; i--) {
        if (i == splitRound) {
          shuffle(seedList, 2**(i-1), lastWithBye - 1);
          shuffle(seedList, lastWithBye, 2**i - 1);
        } else {
          shuffle(seedList, 2**(i-1), 2**i - 1);
        }
      }
      break;
  }
  
  function placeSeeds(seed, location, roundsLeft) {    
    if (2**nround + 1 - seed <= Number(options[2][1])) {
      if (seedMethod != 'random') {bracket.getRange(location, 1).setValue(seedList[seed-1]);}
      bracket.getRange(location, 2).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[seed-1] + 1) + ")");
      if (seedMethod != 'random') {bracket.getRange(location + 1, 1).setValue(seedList[2**nround - seed]);}
      bracket.getRange(location + 1, 2).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[2**nround - seed] + 1) + ")");
    } else {
      bracket.getRange(location, 2, 2, 2).setBorder(false, false, false, false, null, null);
      bracket.getRange(location-1,2).setValue('');
      if ((location - 2) % 8) {
        if (seedMethod != 'random') {bracket.getRange(location - 1, 3).setValue(seedList[seed-1]);}
        bracket.getRange(location - 1, 4).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[seed-1] + 1) + ")");        
      } else {
        if (seedMethod != 'random') {bracket.getRange(location + 2, 3).setValue(seedList[seed-1]);}
        bracket.getRange(location + 2, 4).setFormula("=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (seedList[seed-1] + 1) + ")");        
      }      
    }
    
    for(let i=nround-roundsLeft+1; i < nround; i++) {
      placeSeeds(2**i + 1 - seed, location + 2**(nround+1-i) , nround - i);
    }
    
  }
  placeSeeds(1,2,nround);
  
  //losers' bracket
  bracket.getRange(2**nround + 2, 2*nround + 2, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  bracket.getRange(2**nround + 3, 2*nround, 1, 2).merge().setHorizontalAlignment('right').setValues([['W LF','']]);
  if (options[5][1] != 'grand') {bracket.getRange(2**nround + 2, 2*nround + 4, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
  var ltop = 2**(nround + 1) + 2;
  bracket.getRange('A' + (ltop-1) + ':' + (ltop-1)).setBackground('#707070').setFontColor('#707070');
  
  if (nlround % 2) {
    for (let i = 2**((nlround-1)/2); i>0; i--) {
      bracket.getRange(ltop-3+4*i, 2, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    for (let j = (nlround-3)/2; j>=0; j--) {
      let nshapes = 2**((nlround-3)/2 - j);
      let topShape = 4 * (2**j - 1) - 2*j - 1;
      let spacing = 8 * 2**j;
      for (let i=0; i<nshapes; i++) {
        bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+6, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+4, 1, 2).merge().setHorizontalAlignment('right').setValues([['L ' + (nround - (nlround-3)/2 + j) + '|' + ((j%2) ? nshapes - i : i + 1),'']]);
        bracket.getRange(ltop + topShape + i*spacing + 4, 4*j+4, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        if (j != 0) {bracket.getRange(ltop + topShape + i*spacing + 4 - spacing/4, 4*j+4, spacing/2, 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
        if (nshapes == 1) {bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+6).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue('Match LF');}
      }
    }
    let searchRange = bracket.getRange('B1:B' + 2**(nround+1)).getValues();
    let nmatches = 2**(nround-2)
    for (let i=nmatches; i>0; i--) {
      if (searchRange[8*i-3][0]) {
        bracket.getRange(ltop - 3 + 4*i ,1,2,1).setValues([['L 2|' + (nmatches-i+1)],['L 1|' + (2*i)]]);
      } else {
        bracket.getRange(ltop - 3 + 4*i ,2,2,2).setBorder(false, false, false, false, null, null);
        bracket.getRange(ltop + 4*i - ((i%2) ? 1 : 4), 2, 1, 2).merge().setHorizontalAlignment('right').setValues([['L 2|' + (nmatches-i+1),'']]);
      }
    }
  } else {
    for (let j = (nlround/2) - 1; j>=0; j--) {
      let nshapes = 2**((nlround/2)-j-1);
      let topShape = 3 * (2**j - 1) - 2*j;
      let spacing = 6 * 2**j;
      for (let i=0; i<nshapes; i++) {
        bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+4, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+2, 1, 2).merge().setHorizontalAlignment('right').setValues([['L ' + (nround - (nlround/2) + 1 + j) + '|' + ((j%2) ? i + 1 : nshapes - i), '']]);
        bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+2, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        if (j != 0) {bracket.getRange(ltop + topShape + i*spacing + 3 - spacing/4, 4*j+2, spacing/2, 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
        if (nshapes == 1) {bracket.getRange(ltop + topShape + i*spacing, 4*j+4).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue('Match LF');}
      }
    }
    let searchRange = bracket.getRange('B1:B' + 2**(nround+1)).getValues();
    for (let i=2**(nround-2); i>0; i--) {
      if (searchRange[8*i-7][0]) {
        bracket.getRange(ltop - 3 + 6*i, 1, 2, 1).setValues([['L 1|' + (2*i-1)],['L 1|' + (2*i)]]);
      } else {
        bracket.getRange(ltop - 3 + 6*i, 2, 2, 2).setBorder(false, false, false, false, null, null);
        bracket.getRange(ltop - 4 + 6*i, 2, 1, 2).merge().setHorizontalAlignment('right').setValues([['L 1|' + (2*i),'']]);
      }
    }
  }

  //places sheet if desired
  if (options[6][1]) {
    var places = sheet.insertSheet(name + ' Places');
    places.deleteColumns(3,24);
    places.setColumnWidth(1,40);
    places.setColumnWidth(2,150);
    places.getRange('A:A').setVerticalAlignment('top').setHorizontalAlignment('right');
    places.getRange(1,1,5,2).setValues([
      ['Rank', 'Player'],
      [1, ''],
      [2, ''],
      [3, ''],
      [4, '']
    ]);
    for (let n=2; n<Number(options[2][1])/2; n*=2) {
      places.getRange(2+2*n,1).setValue('T' + (1+2*n));
      places.getRange(2+2*n,1,n,1).merge();
      if (3*n < Number(options[2][1])) {
        places.getRange(2+3*n,1).setValue('T' + (1+3*n));
        places.getRange(2+3*n,1,n,1).merge();
      }
    }
    places.deleteRows(Number(options[2][1]) + 2, 999-Number(options[2][1]));
  }
}

function updateDoubleBracket(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var bracket = sheet.getSheetByName(name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var nlround = 2*nround - 2 - (2**nround-options[2][1]>=2**(nround-2));
  var gamesToWin = options[3][1];
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
  var found = false;
  
  var ltop = 2**(nround + 1) + 2;
  var loserInfo = bracket.getRange(ltop - 1, 1, 1, 2).getValues()[0];
  
  if (loserInfo[1]) {
    //in finals
    if (options[5][1] != 'grand') {
      var match = bracket.getRange(2**nround + 2, 2*nround + 4, 2, 2).getValues();
      if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
        var topwins = Number(match[0][1]) + Number(mostRecent[1]);
        var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
        found = true;
      }
      if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
        var topwins = Number(match[0][1]) + Number(mostRecent[3]);
        var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);      
        found = true;
      }
    }
    if (found) {
      bracket.getRange(2**nround + 2, 2*nround + 5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      var loc = 'f,f';
      if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
        bracket.getRange(2**nround + 2, 2*nround + 4, 1, 2).setBackground('#D9EBD3');
        if (options[6][1]) {sheet.getSheetByName(name + ' Places').getRange(2,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
      } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
        bracket.getRange(2**nround + 3, 2*nround + 4, 1, 2).setBackground('#D9EBD3');
        if (options[6][1]) {sheet.getSheetByName(name + ' Places').getRange(2,2,2,1).setValues([[match[1][0]],[match[0][0]]]);}
      }
    } else {
      var match = bracket.getRange(2**nround + 2, 2*nround + 2, 2, 2).getValues();
      if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
        var topwins = Number(match[0][1]) + Number(mostRecent[1]);
        var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
        found = true;
      }
      if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
        var topwins = Number(match[0][1]) + Number(mostRecent[3]);
        var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
        found = true;
      }
      if (found) {
        bracket.getRange(2**nround + 2, 2*nround + 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
        var loc = 'f,b';
        if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
          bracket.getRange(2**nround + 2, 2*nround + 2, 1, 4).setBackground('#D9EBD3');
          if (options[5][1] != 'grand') {
            bracket.getRange(2**nround + 2, 2*nround + 4).setValue(match[0][0]);
          }
          if (options[6][1]) {sheet.getSheetByName(name + ' Places').getRange(2,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
        } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
          bracket.getRange(2**nround + 3, 2*nround + 2, 1, 2).setBackground('#D9EBD3');
          if (options[5][1] != 'grand') {
            bracket.getRange(2**nround + 2, 2*nround + 4, 2, 1).setValues([[match[0][0]], [match[1][0]]]);
          } else if (options[6][1]) {
            sheet.getSheetByName(name + ' Places').getRange(2,2,2,1).setValues([[match[1][0]],[match[0][0]]]);
          }
        }
      }
    }
  } else if (loserInfo[0].indexOf('~' + mostRecent[0] + '~') == -1) {
    //in winner bracket
    for (let j=nround; j>0; j--) {
      let nmatches = 2**(nround-j);
      let space = 2**j;
      for (let i=0; i<nmatches; i++) {
        if ((bracket.getRange(space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).getValue() === '') || j == nround) {
          let match = bracket.getRange(space + 2*i*space, 2*j, 2, 2).getValues();
          if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
            var topwins = Number(match[0][1]) + Number(mostRecent[1]);
            var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
            found = true;
          } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
            var topwins = Number(match[0][1]) + Number(mostRecent[3]);
            var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
            found = true;
          }
          if (found) {
            bracket.getRange(space + 2*i*space, 2*j+1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            var loc = 'w,' + j + ',' + i;
            if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
              if (j != nround) {
                bracket.getRange(space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[0][0]);
              } else {
                bracket.getRange(space + 2*i*space + 2, 2*j+2).setValue(match[0][0]);
              }
              bracket.getRange(space + 2*i*space, 2*j, 1, 2).setBackground('#D9EBD3');
              bracket.getRange(ltop - 1, 1).setValue(loserInfo[0] + '~' + match[1][0] + '~');
              if (nlround % 2) {
                if (j>2) {
                  bracket.getRange(ltop + 4 * (2**(j-3) - 1) - 2*j + 7 + ((j%2) ? i : -1-i+nmatches)*8*2**(j-3), 4*j-6).setValue(match[1][0]);
                } else if (j == 2) {
                  if (bracket.getRange(ltop + 4*(nmatches - i) - 3,1).getValue()) {
                    bracket.getRange(ltop + 4*(nmatches - i) - 3,2).setValue(match[1][0]);
                  } else {
                    bracket.getRange(ltop + 4*(nmatches - i) - 3 + ((i%2) ? 2 : -1), 4).setValue(match[1][0]);
                  }
                } else {
                  bracket.getRange(ltop + 2*i, 2).setValue(match[1][0]);
                }
              } else {
                if (j>1) {
                  bracket.getRange(ltop + 3 * (2**(j-2) - 1) - 2*j + 5 + ((j%2) ? i : -1-i+2**(nround-j))*6*2**(j-2), 4*j-4).setValue(match[1][0]);
                } else {
                  if (i%2) {
                    if (bracket.getRange(ltop + 1 + 3*i, 1).getValue()) {
                      bracket.getRange(ltop + 1 + 3*i, 2).setValue(match[1][0]);
                    } else {
                      bracket.getRange(ltop - 1 + 3*i, 4).setValue(match[1][0]);
                    }
                  } else {
                    bracket.getRange(ltop + 3 + 3*i, 2).setValue(match[1][0]);
                  }
                }
              }
            } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
              if (j != nround) {
                bracket.getRange(space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[1][0]);
              } else {
                bracket.getRange(space + 2*i*space + 2, 2*j+2).setValue(match[1][0]);
              }
              bracket.getRange(space + 2*i*space+1, 2*j, 1, 2).setBackground('#D9EBD3');
              bracket.getRange(ltop - 1, 1).setValue(loserInfo[0] + '~' + match[0][0] + '~');
              if (nlround % 2) {
                if (j>2) {
                  bracket.getRange(ltop + 4 * (2**(j-3) - 1) - 2*j + 7 + ((j%2) ? i : -1-i+2**(nround-j))*8*2**(j-3), 4*j-6).setValue(match[0][0]);
                } else if (j == 2) {
                  if (bracket.getRange(ltop + 4*(nmatches - i) - 3,1).getValue()) {
                    bracket.getRange(ltop + 4*(nmatches - i) - 3,2).setValue(match[0][0]);
                  } else {
                    bracket.getRange(ltop + 4*(nmatches - i) - 3 + ((i%2) ? 2 : -1), 4).setValue(match[0][0]);
                  }
                } else {
                  bracket.getRange(ltop + 2*i, 2).setValue(match[0][0]);
                }
              } else {
                if (j>1) {
                  bracket.getRange(ltop + 3 * (2**(j-2) - 1) - 2*j + 5 + ((j%2) ? i : -1-i+2**(nround-j))*6*2**(j-2), 4*j-4).setValue(match[0][0]);
                } else {
                  if (i%2) {
                    if (bracket.getRange(ltop + 1 + 3*i, 1).getValue()) {
                      bracket.getRange(ltop + 1 + 3*i, 2).setValue(match[0][0]);
                    } else {
                      bracket.getRange(ltop - 1 + 3*i, 4).setValue(match[0][0]);
                    }
                  } else {
                    bracket.getRange(ltop + 3 + 3*i, 2).setValue(match[0][0]);
                  }
                }
              }
            }
            break;
          }
        }
      }
      if (found) {break;}
    }    
  } else {
    //in loser bracket
    if (nlround % 2) {
      for (let j = (nlround-3)/2; j>=0; j--) {
        let nshapes = 2**((nlround-3)/2 - j);
        let topShape = 4 * (2**j - 1) - 2*j - 1;
        let spacing = 8 * 2**j;
        for (let i=0; i<nshapes; i++) {
          if (bracket.getRange(ltop + topShape + i*spacing + 2 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+8).getValue() === '') {
            let match = bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+6, 2, 2).getValues();
            if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[1]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
              found = true;
            } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[3]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
              found = true;
            }
            if (found) {
              bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
              var loc = 'l,' + j + ',' + i + ',f';
              if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+6, 1, 2).setBackground('#D9EBD3');
                if (j == (nlround-3)/2) {
                  bracket.getRange(2**nround+3, 2*nround+2).setValue(match[0][0]);
                  bracket.getRange(ltop-1, 2).setValue('Finals');
                } else {
                  bracket.getRange(ltop + topShape + i*spacing + 2 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+8).setValue(match[0][0]);
                }
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
                  for (let k=0; k<nshapes; k++) {
                    if(!tofill[k]) {
                      places.getRange(2+k+2*nshapes,2).setValue(match[1][0]);
                      break;
                    }
                  }
                }
              } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+6, 1, 2).setBackground('#D9EBD3');
                if (j == (nlround-3)/2) {
                  bracket.getRange(2**nround+3, 2*nround+2).setValue(match[1][0]);
                  bracket.getRange(ltop-1, 2).setValue('Finals');
                } else {
                  bracket.getRange(ltop + topShape + i*spacing + 2 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+8).setValue(match[1][0]);
                }
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
                  for (let k=0; k<nshapes; k++) {
                    if(!tofill[k]) {
                      places.getRange(2+k+2*nshapes,2).setValue(match[0][0]);
                      break;
                    }
                  }
                }
              }
              break;
            } else if (bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+6).getValue() === '') {
              match = bracket.getRange(ltop + topShape + i*spacing + 4, 4*j+4, 2, 2).getValues();
              var loc = 'l,' + j + ',' + i + ',b';
              if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
                var topwins = Number(match[0][1]) + Number(mostRecent[1]);
                var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
                found = true;
              } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
                var topwins = Number(match[0][1]) + Number(mostRecent[3]);
                var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
                found = true;
              }
              if (found) {
                bracket.getRange(ltop + topShape + i*spacing + 4, 4*j+5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
                if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                  bracket.getRange(ltop + topShape + i*spacing + 4, 4*j+4, 1, 2).setBackground('#D9EBD3');
                  bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+6).setValue(match[0][0]);
                  if (options[6][1]) {
                    let places = sheet.getSheetByName(name + ' Places');
                    let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
                    for (let k=0; k<nshapes; k++) {
                      if(!tofill[k]) {
                        places.getRange(2+k+3*nshapes,2).setValue(match[1][0]);
                        break;
                      }
                    }
                  }
                } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                  bracket.getRange(ltop + topShape + i*spacing + 5, 4*j+4, 1, 2).setBackground('#D9EBD3');
                  bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+6).setValue(match[1][0]);
                  if (options[6][1]) {
                    let places = sheet.getSheetByName(name + ' Places');
                    let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
                    for (let k=0; k<nshapes; k++) {
                      if(!tofill[k]) {
                        places.getRange(2+k+3*nshapes,2).setValue(match[0][0]);
                        break;
                      }
                    }
                  }
                }
                break;
              }
            }
          }
        }
        if (found) {break;}
      }
      if (!found) {
        for (let i=2**(nround-2); i>0; i--) {
          if (bracket.getRange(ltop + 4*i - 3 + ((i%2) ? 2 : -1), 4).getValue() === '') {
            let match = bracket.getRange(ltop + 4*i - 3, 2, 2, 2).getValues();
            if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[1]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
              found = true;
            } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[3]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
              found = true;
            }
            if (found) {
              bracket.getRange(ltop + 4*i - 3, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
              var loc = 'l,f,' + i;
              if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                bracket.getRange(ltop + 4*i - 3, 2, 1, 2).setBackground('#D9EBD3');
                bracket.getRange(ltop + 4*i - 3 + ((i%2) ? 2 : -1), 4).setValue(match[0][0]);
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2**(nround-1),2,2**(nround-2),1).getValues().flat();
                  for (let k=0; k<tofill.length; k++) {
                    Logger.log(tofill);
                    if (!tofill[k]) {
                      places.getRange(2+k+2**(nround-1),2).setValue(match[1][0]);
                      break;
                    }
                  }
                }
              } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                bracket.getRange(ltop + 4*i - 2, 2, 1, 2).setBackground('#D9EBD3');
                bracket.getRange(ltop + 4*i - 3 + ((i%2) ? 2 : -1), 4).setValue(match[1][0]);
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2**(nround-1),2,2**(nround-2),1).getValues().flat();
                  for (let k=0; k<tofill.length; k++) {
                    if (!tofill[k]) {
                      places.getRange(2+k+2**(nround-1),2).setValue(match[0][0]);
                      break;
                    }
                  }
                }
              }
              break;
            }
          }
        }
      }
    } else {
      for (let j = (nlround/2) - 1; j>=0; j--) {
        let nshapes = 2**((nlround/2)-j-1);
        let topShape = 3 * (2**j - 1) - 2*j;
        let spacing = 6 * 2**j;
        for (let i=0; i<nshapes; i++) {
          if (bracket.getRange(ltop + topShape + i*spacing + 1 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+6).getValue() === '') {
            let match = bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+4, 2, 2).getValues();
            if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[1]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
              found = true;
            } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
              var topwins = Number(match[0][1]) + Number(mostRecent[3]);
              var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
              found = true;
            }
            if (found) {
              bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
              var loc = 'l,' + j + ',' + i + ',f';
              if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                bracket.getRange(ltop + topShape + i*spacing + 1, 4*j+4, 1, 2).setBackground('#D9EBD3');
                if (j == (nlround/2) - 1) {
                  bracket.getRange(2**nround+3, 2*nround+2).setValue(match[0][0]);
                  bracket.getRange(ltop-1, 2).setValue('Finals');
                } else {
                  bracket.getRange(ltop + topShape + i*spacing + 1 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+6).setValue(match[0][0]);
                }
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
                  for (let k=0; k<nshapes; k++) {
                    if(!tofill[k]) {
                      places.getRange(2+k+2*nshapes,2).setValue(match[1][0]);
                      break;
                    }
                  }
                }
              } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+4, 1, 2).setBackground('#D9EBD3');
                if (j == (nlround/2) - 1) {
                  bracket.getRange(2**nround+3, 2*nround+2).setValue(match[1][0]);
                  bracket.getRange(ltop-1, 2).setValue('Finals');
                } else {
                  bracket.getRange(ltop + topShape + i*spacing + 1 + ((i%2) ? 1 - spacing/2 : spacing/2), 4*j+6).setValue(match[1][0]);
                }
                if (options[6][1]) {
                  let places = sheet.getSheetByName(name + ' Places');
                  let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
                  for (let k=0; k<nshapes; k++) {
                    if(!tofill[k]) {
                      places.getRange(2+k+2*nshapes,2).setValue(match[0][0]);
                      break;
                    }
                  }
                }
              }
              break;
            } else if (bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+4).getValue() === '') {
              match = bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+2, 2, 2).getValues();
              if ((match[0][0] == mostRecent[0]) && (match[1][0] == mostRecent[2])) {
                var topwins = Number(match[0][1]) + Number(mostRecent[1]);
                var bottomwins = Number(match[1][1]) + Number(mostRecent[3]);
                found = true;
              } else if ((match[0][0] == mostRecent[2]) && (match[1][0] == mostRecent[0])) {
                var topwins = Number(match[0][1]) + Number(mostRecent[3]);
                var bottomwins = Number(match[1][1]) + Number(mostRecent[1]);
                found = true;
              }
              if (found) {
                bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
                var loc = 'l,' + j + ',' + i + ',b';
                if ((topwins >= gamesToWin) && (topwins > bottomwins)) {
                  bracket.getRange(ltop + topShape + i*spacing + 3, 4*j+2, 1, 2).setBackground('#D9EBD3');
                  bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+4).setValue(match[0][0]);
                  if (options[6][1]) {
                    let places = sheet.getSheetByName(name + ' Places');
                    let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
                    for (let k=0; k<nshapes; k++) {
                      if(!tofill[k]) {
                        places.getRange(2+k+3*nshapes,2).setValue(match[1][0]);
                        break;
                      }
                    }
                  }
                } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                  bracket.getRange(ltop + topShape + i*spacing + 4, 4*j+2, 1, 2).setBackground('#D9EBD3');
                  bracket.getRange(ltop + topShape + i*spacing + 2, 4*j+4).setValue(match[1][0]);
                  if (options[6][1]) {
                    let places = sheet.getSheetByName(name + ' Places');
                    let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
                    for (let k=0; k<nshapes; k++) {
                      if(!tofill[k]) {
                        places.getRange(2+k+3*nshapes,2).setValue(match[0][0]);
                        break;
                      }
                    }
                  }
                }
                break;
              }
            }
          }       
        }
        if (found) {break;}
      }
    }
  }
  
  if (found && !empty) {
    sendResultToDiscord(name, 'Bracket', bracket.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(loc);    
  }
  return found;
}

function removeDoubleBracket(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',');
  var bracket = sheet.getSheetByName(name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var nlround = 2*nround - 2 - (2**nround-options[2][1]>=2**(nround-2));
  var ltop = 2**(nround + 1) + 2;
  var loserInfo = bracket.getRange(ltop - 1, 1, 1, 2).getValues()[0];
  
  if (loc[0] == 'f') {
    //in finals
    if (loc[1] == 'f') {
      var match = bracket.getRange(2**nround+2, 2*nround+4, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      bracket.getRange(2**nround+2, 2*nround+5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
    } else {
      if (bracket.getRange(2**nround+2, 2*nround+5).getValue() !== '') {
        SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The second finals match has a result');
        return;
      } else {
        bracket.getRange(2**nround+2, 2*nround+4, 2, 2).setBackground(null).setValues([['',''],['','']]);
        var match = bracket.getRange(2**nround+2, 2*nround+2, 2, 2).setBackground(null).getValues();
        if (removal[1] == match[0][0]) {
          var topwins = Number(match[0][1]) - Number(removal[2]);
          var bottomwins = Number(match[1][1]) - Number(removal[4]);
        } else {
          var topwins = Number(match[0][1]) - Number(removal[4]);
          var bottomwins = Number(match[1][1]) - Number(removal[2]);
        }
        bracket.getRange(2**nround+2, 2*nround+3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      }
    }
    if (options[6][1]) {sheet.getSheetByName(name + ' Places').getRange(2,2,2,1).setValues([[''],['']]);}
  } else if (loc[0] == 'w') {
    //in winners' bracket
    var space = 2**loc[1];
    var nmatches = 2**(nround-loc[1]);
    //scrub dependencies on this match and fail if one has a result
    var toScrub = [];
    if (bracket.getRange(space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+3).getValue() !== '') {
      SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
      return;
    } else {
      toScrub.push(bracket.getRange(space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+2));
    }
    if (nlround % 2) {
      if (loc[1]>2) {
        if (bracket.getRange(ltop + 4 * (2**(loc[1]-3) - 1) - 2*loc[1] + 7 + ((loc[1]%2) ? loc[2] : -1-loc[2]+nmatches)*8*2**(loc[1]-3), 4*loc[1]-5).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 4 * (2**(loc[1]-3) - 1) - 2*loc[1] + 7 + ((loc[1]%2) ? loc[2] : -1-loc[2]+nmatches)*8*2**(loc[1]-3), 4*loc[1]-6));
        }
      } else if (loc[1] == 2) {
        if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,1).getValue()) {
          if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,3).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
            return;
          } else {
            toScrub.push(bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,2));
          }
        } else {
          if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3 + ((loc[2]%2) ? 2 : -1), 5).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
            return;
          } else {
            toScrub.push(bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3 + ((loc[2]%2) ? 2 : -1), 4));
          }
        }
      } else {
        if (bracket.getRange(ltop + 2*loc[2], 3).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 2*loc[2], 2));
        }
      }
    } else {
      if (loc[1]>1) {
        if (bracket.getRange(ltop + 3 * (2**(loc[1]-2) - 1) - 2*loc[1] + 5 + ((loc[2]%2) ? loc[2] : -1-loc[2]+2**(nround-loc[1]))*6*2**(loc[1]-2), 4*loc[1]-3).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 3 * (2**(loc[1]-2) - 1) - 2*loc[1] + 5 + ((loc[2]%2) ? loc[2] : -1-loc[2]+2**(nround-loc[1]))*6*2**(loc[1]-2), 4*loc[1]-4));
        }
      } else {
        if (loc[2]%2) {
          if (bracket.getRange(ltop + 1 + 3*loc[2], 1).getValue()) {
            if (bracket.getRange(ltop + 1 + 3*loc[2], 3).getValue() !== '') {
              SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
              return;
            } else {
              toScrub.push(bracket.getRange(ltop + 1 + 3*loc[2], 2));
            }
          } else {
            if (bracket.getRange(ltop - 1 + 3*loc[2], 5).getValue() !== '') {
              SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
              return;
            } else {
              toScrub.push(bracket.getRange(ltop - 1 + 3*loc[2], 4));
            }
          }
        } else {
          if (bracket.getRange(ltop + 3 + 3*loc[2], 3).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the loser plays in has a result');
            return;
          } else {
            toScrub.push(bracket.getRange(ltop + 3 + 3*loc[2], 2));
          }
        }
      }
    }
    for (let i=toScrub.length; i>0; i--) {toScrub[i-1].setValue('');}
    
    var match = bracket.getRange(space + 2*space*loc[2], 2*loc[1], 2, 2).setBackground(null).getValues();
    loserInfo[0] = loserInfo[0].split('~' + match[0][0] + '~').join('');
    loserInfo[0] = loserInfo[0].split('~' + match[1][0] + '~').join('');
    bracket.getRange(ltop-1, 1).setValue(loserInfo[0]);
    if (removal[1] == match[0][0]) {
      var topwins = Number(match[0][1]) - Number(removal[2]);
      var bottomwins = Number(match[1][1]) - Number(removal[4]);
    } else {
      var topwins = Number(match[0][1]) - Number(removal[4]);
      var bottomwins = Number(match[1][1]) - Number(removal[2]);
    }
    bracket.getRange(space + 2*space*loc[2], 2*loc[1] + 1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
  } else {
    //in losers' bracket
    if (nlround % 2) {
      if (loc[1] == 'f') {
        if (bracket.getRange(ltop + 4*loc[2] - 3 + ((loc[2]%2) ? 2 : -1), 5).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
          return false;
        } else {
          bracket.getRange(ltop + 4*loc[2] - 3 + ((loc[2]%2) ? 2 : -1), 4).setValue('');
        }
        var match = bracket.getRange(ltop + 4*loc[2] - 3, 2, 2, 2).setBackground(null).getValues();
        if (removal[1] == match[0][0]) {
          var topwins = Number(match[0][1]) - Number(removal[2]);
          var bottomwins = Number(match[1][1]) - Number(removal[4]);
        } else {
          var topwins = Number(match[0][1]) - Number(removal[4]);
          var bottomwins = Number(match[1][1]) - Number(removal[2]);
        }
        bracket.getRange(ltop + 4*loc[2] - 3, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
        if (options[6][1]) {
          let places = sheet.getSheetByName(name + ' Places');
          let tofill = places.getRange(2+2**(nround-1),2,2**(nround-2),1).getValues().flat();
          for (let k=0; k<tofill.length; k++) {
            if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
              places.getRange(2+k+2**(nround-1),2).setValue('');
            }
          }
        }
      } else {
        let nshapes = 2**((nlround-3)/2 - loc[1]);
        let topShape = 4 * (2**loc[1] - 1) - 2*loc[1] - 1;
        let spacing = 8 * 2**loc[1];
        if (loc[3] == 'f') {
          if (loc[1] == (nlround-3)/2) {
            if (bracket.getRange(2**nround+3, 2*nround+3).getValue() !== '') {
              SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
              return false;
            } else {
              bracket.getRange(2**nround+3, 2*nround+2).setValue('');
              bracket.getRange(ltop-1, 2).setValue('');
            }
          } else {
            if (bracket.getRange(ltop + topShape + loc[2]*spacing + 2 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+9).getValue() !== '') {
              SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
              return false;
            } else {
              bracket.getRange(ltop + topShape + loc[2]*spacing + 2 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+8).setValue('');
            }
          }
          var match = bracket.getRange(ltop + topShape + loc[2]*spacing + 2, 4*loc[1]+6, 2, 2).setBackground(null).getValues();
          if (removal[1] == match[0][0]) {
            var topwins = Number(match[0][1]) - Number(removal[2]);
            var bottomwins = Number(match[1][1]) - Number(removal[4]);
          } else {
            var topwins = Number(match[0][1]) - Number(removal[4]);
            var bottomwins = Number(match[1][1]) - Number(removal[2]);
          }
          bracket.getRange(ltop + topShape + loc[2]*spacing + 2, 4*loc[1]+7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
          if (options[6][1]) {
            let places = sheet.getSheetByName(name + ' Places');
            let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+2*nshapes,2).setValue('');
              }
            }
          }
        } else {
          if (bracket.getRange(ltop + topShape + loc[2]*spacing + 3, 4*loc[1]+7).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
            return false;
          } else {
            bracket.getRange(ltop + topShape + loc[2]*spacing + 3, 4*loc[1]+6).setValue('');
          }
          var match = bracket.getRange(ltop + topShape + loc[2]*spacing + 4, 4*loc[1]+4, 2, 2).setBackground(null).getValues();
          if (removal[1] == match[0][0]) {
            var topwins = Number(match[0][1]) - Number(removal[2]);
            var bottomwins = Number(match[1][1]) - Number(removal[4]);
          } else {
            var topwins = Number(match[0][1]) - Number(removal[4]);
            var bottomwins = Number(match[1][1]) - Number(removal[2]);
          }
          bracket.getRange(ltop + topShape + loc[2]*spacing + 4, 4*loc[1]+5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
          if (options[6][1]) {
            let places = sheet.getSheetByName(name + ' Places');
            let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+3*nshapes,2).setValue('');
              }
            }
          }
        }
      }
    } else {      
      let nshapes = 2**((nlround/2)-loc[1]-1);
      let topShape = 3 * (2**loc[1] - 1) - 2*loc[1];
      let spacing = 6 * 2**loc[1];
      if (loc[3] == 'f') {
        if (loc[1] == (nlround/2) - 1) {
          if (bracket.getRange(2**nround+3, 2*nround+3).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
            return false;
          } else {
            bracket.getRange(2**nround+3, 2*nround+2).setValue('');
            bracket.getRange(ltop-1, 2).setValue('');
          }
        } else {
          if (bracket.getRange(ltop + topShape + loc[2]*spacing + 1 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+7).getValue() !== '') {
            SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
            return false;
          } else {
            bracket.getRange(ltop + topShape + loc[2]*spacing + 1 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+6).setValue('');
          }
        }
        var match = bracket.getRange(ltop + topShape + loc[2]*spacing + 1, 4*loc[1]+4, 2, 2).setBackground(null).getValues();
        if (removal[1] == match[0][0]) {
          var topwins = Number(match[0][1]) - Number(removal[2]);
          var bottomwins = Number(match[1][1]) - Number(removal[4]);
        } else {
          var topwins = Number(match[0][1]) - Number(removal[4]);
          var bottomwins = Number(match[1][1]) - Number(removal[2]);
        }
        bracket.getRange(ltop + topShape + loc[2]*spacing + 1, 4*loc[1]+5, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
        if (options[6][1]) {
            let places = sheet.getSheetByName(name + ' Places');
            let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+2*nshapes,2).setValue('');
              }
            }
          }
      } else {
        if (bracket.getRange(ltop + topShape + loc[2]*spacing + 2, 4*loc[1]+5).getValue() !== '') {
          SpreadsheetApp.getUi().alert('Removal Failure (Double Elimination): The match the winner plays in has a result');
          return false;
        } else {
          bracket.getRange(ltop + topShape + loc[2]*spacing + 2, 4*loc[1]+4).setValue('');
        }
        var match = bracket.getRange(ltop + topShape + loc[2]*spacing + 3, 4*loc[1]+2, 2, 2).setBackground(null).getValues();
        if (removal[1] == match[0][0]) {
          var topwins = Number(match[0][1]) - Number(removal[2]);
          var bottomwins = Number(match[1][1]) - Number(removal[4]);
        } else {
          var topwins = Number(match[0][1]) - Number(removal[4]);
          var bottomwins = Number(match[1][1]) - Number(removal[2]);
        }
        bracket.getRange(ltop + topShape + loc[2]*spacing + 3, 4*loc[1]+3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
        if (options[6][1]) {
            let places = sheet.getSheetByName(name + ' Places');
            let tofill = places.getRange(2+3*nshapes,2,nshapes,1).getValues().flat();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+3*nshapes,2).setValue('');
              }
            }
          }
      }
    }
  }
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

// Swiss

function swissOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,7,2).setValues([
    ["''" + name + "' Swiss Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""],
    ["Seeding Method", "standard"],
    ["Bye Result", ""]
  ]);

  options.getRange(base + 6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'highlow'],true).build());
}

function createSwiss(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var swiss = sheet.insertSheet(name + ' Standings');
  swiss.deleteRows(2*Math.ceil(options[2][1]/2) + 1, 1000 - 2*Math.ceil(options[2][1]/2));
  swiss.deleteColumns(10, 17)
  swiss.setColumnWidth(1, 20);
  swiss.setColumnWidth(2, 150)
  swiss.setColumnWidths(3, 2, 50);
  swiss.setColumnWidth(6, 150);
  swiss.setColumnWidths(7, 2, 30);
  swiss.setColumnWidth(9, 150);
  swiss.setFrozenColumns(5);
  swiss.setFrozenRows(1);
  swiss.getRange('F1').setHorizontalAlignment('right');
  swiss.getRange('G1').setHorizontalAlignment('left');
  swiss.getRange('I:I').setHorizontalAlignment('right');
  swiss.getRange('H1:I1').setFontColor('#ffffff');
  swiss.getRange('I:I').setHorizontalAlignment('right');
  swiss.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISBLANK(B1)').setFontColor('#FFFFFF').setRanges([swiss.getRange('A:A')]).build()]);
  var processing = sheet.insertSheet(name + ' Processing');
  processing.getRange('E1').setValue('2');
  var players = Array.from({length:Number(options[2][1])}, (v,i) => "=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (i+2) + ")");
  
  //display sheet
  //standings
  var standings = [['Pl', 'Player', 'Wins', 'Losses']];
  for (let i=1; i <= options[2][1]; i++) {
    standings.push(['=if(and(C' + (i+1) + '=C' + i + ', D' + (i+1) + '=D' + i +'), "", ' + i + ')',
     "='" + name + " Processing'!G" + (i+1), "='" + name + " Processing'!H" + (i+1), "='" + name + " Processing'!I" + (i+1)]);
  }
  swiss.getRange(1, 1, standings.length, 4).setValues(standings);
  
  //divider
  swiss.getRange('E:E').setBackground('#707070');
  
  //first round
  if (options[2][1] % 2) {
    var byeplace = Math.floor(Math.random() * (options[2][1] + 1));
    players.splice(byeplace, 0, 'BYE');
  }
  var halfplayers = players.length/2;
  switch (options[5][1]) {
    case 'random':
      shuffle(players);
      if (byeplace) {byeplace = players.indexOf('BYE');}
      break;
    case 'highlow':
      players = players.slice(0,halfplayers).concat(players.slice(halfplayers,2*halfplayers).reverse());
      if (byeplace >= halfplayers) {byeplace = players.indexOf('BYE', halfplayers);}
      break;
  }
  var walkover = [options[4][1], 0];
  var byediff = 0;
  if (/^\d+--\d+$/.test(options[6][1])) {
    walkover = options[6][1].split("--");
    byediff = Number(options[4][1]) - (Number(walkover[0]) + Number(walkover[1]));
  }
  var firstround = [['Round', 1, '=(SUM(G2:H) + I1)/' + (Number(options[4][1]) * halfplayers), 0]];
  for (let i = 0; i < halfplayers; i++) {
    if (players[i] == 'BYE') {
      firstround.push(['BYE', walkover[1], walkover[0], players[halfplayers+i]]);
      firstround[0][3] += byediff;
    } else if (players[halfplayers + i] == 'BYE') {
      firstround.push([players[i], walkover[0], walkover[1], 'BYE']);
      firstround[0][3] += byediff;
    } else {
      firstround.push([players[i], '', '', players[halfplayers+i]]);
    }
    
  }
  swiss.getRange(1, 6, firstround.length, 4).setValues(firstround);
  
  //processing sheet
  processing.getRange('G1').setValue('=query(A:D, "select A, sum(C), sum(D) where not A = \'BYE\' group by A order by sum(C) desc, sum(D)")');
  if(byeplace) {
    let playerToInsert = (byeplace < halfplayers) ? players[halfplayers + byeplace] : players[byeplace - halfplayers];
    processing.getRange(2, 1, 1, 4).setValues([[playerToInsert, 'BYE', walkover[0], walkover[1]]]);
    processing.getRange(3, 1, 1, 4).setValues([['BYE', playerToInsert, walkover[1], walkover[0]]]);
    processing.getRange('E1').setValue(4);
    processing.getRange('E2').setValue(2);
  }
}

function updateSwiss(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var swiss = sheet.getSheetByName(name + ' Standings');
  var processing = sheet.getSheetByName(name + ' Processing');
  var nextOpen = Number(processing.getRange('E1').getValue());
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  
  //update matches and standings - standings updates only occur if the reported match was in the match list
  var roundMatches = swiss.getRange(2, 6, Math.ceil(options[2][1]/2),4).getValues();
  var forceFinish = /count as complete/.test(mostRecent[4].toLowerCase());
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0) && !forceFinish;
  var found = false;
  for (let i=0; i < roundMatches.length; i++) {
    if ((roundMatches[i][0] == mostRecent[0]) && (roundMatches[i][3] == mostRecent[2])) {
      var leftwins = Number(roundMatches[i][1]) + Number(mostRecent[1]);
      var rightwins = Number(roundMatches[i][2]) + Number(mostRecent[3]);
      found = true;
    } else if ((roundMatches[i][0] == mostRecent[2]) && (roundMatches[i][3] == mostRecent[0])) {
      var leftwins = Number(roundMatches[i][1]) + Number(mostRecent[3]);
      var rightwins = Number(roundMatches[i][2]) + Number(mostRecent[1]);
      found = true;
    }    
    if (found && !empty) {
      swiss.getRange(i+2, 7, 1, 2).setValues([[leftwins, rightwins]]);
      if (forceFinish) {
        let fudges = Number(swiss.getRange('I1').getValue());
        fudges += Number(options[4][1]) - (Number(mostRecent[1]) + Number(mostRecent[3]));
        swiss.getRange('I1').setValue(fudges);
      }
      processing.getRange(nextOpen, 1, 2, 4).setValues([[mostRecent[0], mostRecent[2], mostRecent[1], mostRecent[3]], [mostRecent[2], mostRecent[0], mostRecent[3], mostRecent[1]]]);
      processing.getRange('E1').setValue(nextOpen + 2);
      sendResultToDiscord(name, 'Standings', swiss.getSheetId());
      var results = sheet.getSheetByName('Results');
      var resultsLength = results.getDataRange().getNumRows();
      results.getRange(resultsLength, 7).setValue(name);
      results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(swiss.getRange('G1').getValue() + ',' + i + ',' + nextOpen);
      //make next round's match list if needed
      newSwissRound(num, name);
      break;
    }
  }
  
  return found;
}

function newSwissRound(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var swiss = sheet.getSheetByName(name + ' Standings');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var roundInfo = swiss.getRange('G1:H1').getValues()[0];
  
  if ((roundInfo[1] == 1) && (Number(roundInfo[0]) < Number(options[3][1]))) {
    var players = swiss.getRange(2, 2, options[2][1], 1).getValues().flat();
    var assigned = [];
    var processing = sheet.getSheetByName(name + ' Processing');
    var playedMatches = processing.getRange(1,1,processing.getLastRow(),2).getValues();
    
    //trim playedMatches
    var pmlen = playedMatches.length;
    for (let i=pmlen-1; i >= 0; i--) {
      if (!playedMatches[i][0]) {
        playedMatches.splice(i,1);
        break;
      }
    }
    pmlen = playedMatches.length;
    
    //add potential bye
    if (players.length % 2) {
      var byeplace = Math.floor(Math.random() * (options[2][1] + 1));
      players.splice(byeplace, 0, 'BYE');
    }
    
    var halfplayers = players.length/2;
    
    function alreadyPlayed(player) {
      let playedList = [];
      for (let i=0; i<pmlen; i++) {
        if (playedMatches[i][0] == player) {
          playedList.push(playedMatches[i][1]);
        }
      }
      return playedList;
    }

    function searchForNew(player, index) {
      let played = alreadyPlayed(player);

      for (let i=index; i<2*halfplayers; i++) {
        if (!played.includes(players[i]) && !assigned.includes(players[i])) {
          return i;
        }
      }
      
      return -1;
    }
    
    for (let i=0; i<2*halfplayers; i++) {
      if (!assigned.includes(players[i])) {
        let oppIdx = searchForNew(players[i], i+1);
        if (oppIdx >= 0) {
          assigned.push(players[i]);
          assigned.push(players[oppIdx]);
        } else {
          let iplayed = alreadyPlayed(players[i]);
          for(let j=assigned.length - 1; j>=0; j--) {
            if (!iplayed.includes(assigned[j])) {
              if (j % 2) {
                let jNewIdx = searchForNew(assigned[j-1], i+1);
                if (jNewIdx >= 0) {
                  assigned.push(assigned[j-1]);
                  assigned.push(players[jNewIdx]);
                  assigned[j-1] = players[i];
                  break;
                }
              } else {
                let jNewIdx = searchForNew(assigned[j+1], i+1);
                if (jNewIdx >= 0) {
                  assigned.push(assigned[j+1]);
                  assigned.push(players[jNewIdx]);
                  assigned[j+1] = players[i];
                  break;
                }
              }
            }
          }
        }
      }
    }

    var walkover = [options[4][1], 0];
    var byediff = 0;
    if (/^\d+--\d+$/.test(options[6][1])) {
      walkover = options[6][1].split("--");
      byediff = Number(options[4][1]) - (Number(walkover[0]) + Number(walkover[1]));
    }
    var roundMatches = [['Round', Number(roundInfo[0]) + 1, '=(SUM(G2:H) + I1)/' + (Number(options[4][1]) * halfplayers), 0]];
    for (let i=0; i<halfplayers; i++) {
      if (assigned[2*i] == 'BYE') {
        roundMatches.push(['BYE', walkover[1], walkover[0], assigned[2*i+1]]);
        roundMatches[0][3] += byediff;
      } else if (assigned[2*i+1] == 'BYE') {
        roundMatches.push([assigned[2*i], walkover[0], walkover[1], 'BYE']);
        roundMatches[0][3] += byediff;
      } else {
        roundMatches.push([assigned[2*i], '', '', assigned[2*i+1]]);
      }
    }
    
    swiss.insertColumns(6, 5);
    swiss.setColumnWidth(6, 150);
    swiss.setColumnWidths(7, 2, 30);
    swiss.setColumnWidth(9, 150);
    swiss.setColumnWidth(10, 100);
    swiss.getRange('F1').setHorizontalAlignment('right');
    swiss.getRange('G1').setHorizontalAlignment('left');
    swiss.getRange('H1:I1').setFontColor('#ffffff');
    swiss.getRange('I:I').setHorizontalAlignment('right');
    swiss.getRange('J:J').setBackground('#707070');
    swiss.getRange(1, 6, roundMatches.length, 4).setValues(roundMatches);
    
    if (byeplace) {
      let byeIdx = assigned.indexOf('BYE');
      let nextOpen = Number(processing.getRange('E1').getValue());      
      if (byeIdx % 2) {
        processing.getRange(nextOpen, 1, 2, 4).setValues([[assigned[byeIdx-1],'BYE', walkover[0], walkover[1]],['BYE', assigned[byeIdx-1], walkover[1], walkover[0]]]);
      } else {
        processing.getRange(nextOpen, 1, 2, 4).setValues([[assigned[byeIdx+1],'BYE', walkover[0], walkover[1]],['BYE', assigned[byeIdx+1], walkover[1], walkover[0]]]);
      }
      processing.getRange('E1').setValue(nextOpen + 2);
      processing.getRange(Number(roundInfo[0]+2), 5).setValue(nextOpen);
    }
  }
}

function removeSwiss(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',').map(x=>Number(x));
  var swiss = sheet.getSheetByName(name + ' Standings');
  var round = Number(swiss.getRange('G1').getValue());
  var processing = sheet.getSheetByName(name + ' Processing');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var forceFinish = /count as complete/.test(removal[5].toLowerCase());
  
  if (loc[0] + 1 < round) {
    Logger.log('Removal Failure (Swiss): The next round has started');
    return false;
  } else if (loc[0] + 1 == round) {
    let roundMatches = swiss.getRange(2, 6, Math.ceil(options[2][1]/2), 4).getValues();
    for (let i = roundMatches.length - 1; i>=0; i--) {
      if ((roundMatches[i][1] !== '') && (roundMatches[i][0] != 'BYE') && (roundMatches[i][3] != 'BYE')) {
        Logger.log('Removal Failure (Swiss): The next round has started');
        return false;
      }
    }
    swiss.deleteColumns(6, 5);
    if (options[2][1] % 2) {
      processing.getRange(processing.getRange(round+1,5).getValue(), 1, 2, 4).setValues([['','','',''],['','','','']]);
    }
  }
  processing.getRange(loc[2], 1, 2, 4).setValues([['','','',''],['','','','']]);
  let match = swiss.getRange(loc[1] + 2, 6, 1, 4).getValues()[0];
  if (match[0] == removal[1]) {
    var leftwins = match[1] - removal[2];
    var rightwins = match[2] - removal[4];
  } else {
    var leftwins = match[1] - removal[4];
    var rightwins = match[2] - removal[2];
  }
  swiss.getRange(loc[1] + 2, 7, 1, 2).setValues(((leftwins == 0) && (rightwins == 0)) ? [['','']] : [[leftwins, rightwins]]);
  if (forceFinish) {
    let fudges = Number(swiss.getRange('I1').getValue());
    fudges -= Number(options[4][1]) - (Number(removal[2]) + Number(removal[4]));
    swiss.getRange('I1').setValue(fudges);
  }
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

// Random

function randomOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,5,2).setValues([
    ["''" + name + "' Preset Random Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""]
  ]);
}

function createRandom(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var playerNames = sheet.getSheetByName(options[1][1]).getRange("B2:B" + (Number(options[2][1]) + 1)).getValues().flat();
  
  //setup matches
  var standings = sheet.insertSheet(name + ' Standings');
  standings.deleteRows(Number(options[2][1])+2,999-Number(options[2][1]));
  standings.deleteColumns(8, 19);
  standings.insertColumnsAfter(7, 2*options[3][1]+1)
  standings.setColumnWidth(1, 20);
  standings.setColumnWidth(2, 150);
  standings.setColumnWidths(3, 4, 50);
  standings.getRange('C:C').setNumberFormat('0.000');
  standings.getRange('F:F').setNumberFormat('0.000');
  standings.getRange('G:G').setBackground('#707070');
  standings.setColumnWidth(8, 150);
  for (let i=0; i<options[3][1]; i++) {
    standings.setColumnWidth(9+2*i, 150);
    standings.getRange(1, 9+2*i, options[2][1]+1).setBorder(false, true, false, false, null, false);
    standings.getRange(2, 9+2*i, options[2][1]).setFontColor('#990000');
    standings.setColumnWidth(10+2*i, 20);
  }
  standings.setFrozenColumns(7);
  standings.setFrozenRows(1);
  standings.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISBLANK(B1)').setFontColor('#FFFFFF').setRanges([standings.getRange('A:A')]).build()]);
  var players = Array.from({length:Number(options[2][1])}, (v,i) => ["=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (i+2) + ")", playerNames[i]]);
  players.sort((a,b) => lowerSort(a[1], b[1]));
  players = players.map(x => x[0]);

  if (options[2][1] % 2) {players.push(['OPEN']);}
  var pllength = players.length;
  
  var matches = [];
  for (let i=0; i<pllength; i++) {
    matches.push([players[i]]);
  }
  
  function matchExists(p1,p2) {
    for (let j=0; j<pllength; j++) {
      if (matches[j][0] == p1) {
        for (let k=0; k<options[3][1]; k++) {
          if (matches[j][2*k+1] == p2) {
            return true;
          }
        }
      }
    }
    return false;
  }
  
  //assign matches
  for (let rnd=0; rnd<options[3][1]; rnd++) {
    let bad = false;
    let potential = [];
    do {
      potential = [];
      bad = false;
      shuffle(players);
      for (let i=0; i<options[2][1]/2; i++) {
        if (rnd) {
          bad = matchExists(players[2*i], players[2*i+1]);
        }
        if (!bad) {
          potential.push([players[2*i], players[2*i+1]]);
          potential.push([players[2*i+1], players[2*i]]);
        } else {break;}
      }
    } while (bad);
    for (let i=0; i<pllength; i++) {
      for (let j=0; j<pllength; j++) {
        if (potential[i][0] == matches[j][0]) {
          matches[j].push(potential[i][1]);
          matches[j].push('');
        }
      }
    }
  }
  
  //handle odd number of players
  var byePlayer = null;
  if (options[2][1] % 2) {
    let opens = [];
    for (let i=0; i<pllength; i++) {
      if (matches[i][0] == 'OPEN') {
        for (let rnd=0; rnd<options[3][1]; rnd++) {
          opens.push(matches[i][rnd*2+1]);
        }
        matches.splice(i,1);
        pllength--;
        break;
      }
    }
    if (opens.length % 2) {opens.splice(Math.floor(Math.random() * opens.length), 0, 'BYE');}
    let assigned = [];
    for (let i=0; i<opens.length; i++) {
      if (!assigned.includes(opens[i])) {
        let found = false;
        for (let j=i+1; j<opens.length; j++) {
          if (!matchExists(opens[i], opens[j])) {
            assigned.push(opens[i]);
            assigned.push(opens[j]);
            found = true;
            break;
          }
        }
        if (!found) {
          for (let j=assigned.length-1; j>=0; j--) {
            if (!matchExists(opens[i], assigned[j])) {
              if (j%2) {
                for (let k=i+1; k<opens.length; k++) {
                  if (!matchExists(assigned[j-1], opens[k])) {
                    assigned.push(assigned[j-1]);
                    assigned.push(opens[k]);
                    assigned[j-1] = opens[i];
                    found = true;
                    break;
                  }
                }
              } else {
                for (let k=i+1; k<opens.length; k++) {
                  if (!matchExists(assigned[j+1], opens[k])) {
                    assigned.push(assigned[j+1]);
                    assigned.push(opens[k]);
                    assigned[j-1] = opens[i];
                    found = true;
                    break;
                  }
                }
              }
            }
            if (found) {break;}
          }
        }
        if (!found) {
          //this is an exceedingly rare case where someone cannot be found an unplayed opponent due to previous randomization
          //for ease, just try again from scratch
          sheet.deleteSheet(standings);
          createRandom(num, name);
          return
        }
      }
    }
    for (let i=0; i<assigned.length; i++) {
      if (assigned[i] == 'BYE') {
        byePlayer = (i % 2) ? assigned[i-1] : assigned[i+1]
      } else {
        for (let j=0; j<pllength; j++) {
          if (matches[j][0] == assigned[i]) {
            for (let k=0; k<options[3][1]; k++) {
              if (matches[j][2*k+1] == 'OPEN') {
                let newval = (i % 2) ? assigned[i-1] : assigned[i+1]
                matches[j][2*k+1] = newval;
                if (newval == 'BYE') {
                  matches[j][2*k+2] = options[4][1];
                  standings.getRange(j+2, 9+2*k).setFontColor('#38761d');
                }
              }
            }
          }
        }
      }
    }
  }
  
  //put matches on sheet

  var topRow = ['Player'];
  for (let i=0; i<options[3][1]; i++) {
    topRow.push('Opponent ' + (i+1));
    topRow.push('');
  }
  matches.splice(0,0,topRow);
  standings.getRange(1, 8, matches.length, matches[0].length).setValues(matches).applyRowBanding();
  
  //results processing
  var processing = sheet.insertSheet(name + ' Processing');
  processing.getRange('D1').setValue(1);
  processing.getRange('E1').setFormula('=query(K:N,"select K,L/M,L,M-L,N where not K = \'BYE\' order by L/M desc, N desc, L desc")');
  processing.getRange('K1').setFormula('=query(A:C,"select A, sum(B), sum(B)+sum(C) group by A order by A")');
  var sosFormulas = [];
  for (let i=0; i<pllength+Boolean(byePlayer); i++) {
    let rowformulas = ["=IFERROR((O"+(i+3)+"+L"+(i+3)+"-M"+(i+3)+")/(P"+(i+3)+"-M"+(i+3)+"),\"\")", "=", "="];
    for (let j=0; j<options[3][1]; j++) {
      rowformulas[1] += "+".repeat(j!=0) + "IFERROR(VLOOKUP(VLOOKUP(K" + (i+3) + ",'" + name + " Standings'!H2:" + (options[2][1]+1) + "," + (2*j+2) + "),K:M,2,FALSE),0)";
      rowformulas[2] += "+".repeat(j!=0) + "IFERROR(VLOOKUP(VLOOKUP(K" + (i+3) + ",'" + name + " Standings'!H2:" + (options[2][1]+1) + "," + (2*j+2) + "),K:M,3,FALSE),0)";
    }
    sosFormulas.push(rowformulas);
  }
  processing.getRange(3, 14, sosFormulas.length, 3).setValues(sosFormulas);
  if (byePlayer) {
    processing.getRange(1, 1, 2, 3).setValues([[byePlayer, options[4][1], 0], ['BYE', 0, options[4][1]]]);
    processing.getRange('D1').setValue(3);
  }
  
  //standings table copy from processing
  var table = [['Pl', 'Player', 'Win Pct','Wins', 'Losses', 'SoS']];
  for (let i=1; i <= options[2][1]; i++) {
    table.push(['=if(and(C' + (i+1) + '=C' + i + ', F' + (i+1) + '=F' + i +'), "", ' + i + ')',
     "='"+name+" Processing'!E"+(i+1), "='"+name+" Processing'!F"+(i+1),
      "='"+name+" Processing'!G"+(i+1), "='"+name+" Processing'!H"+(i+1), "='"+name+" Processing'!I"+(i+1)]);
  }
  standings.getRange(1,1,table.length,6).setValues(table);
}

function updateRandom(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var standings = sheet.getSheetByName(name + ' Standings');
  var matches = standings.getRange(2, 8, options[2][1], 2*options[3][1]+1).getValues();
  var mlength = matches.length;
  var processing = sheet.getSheetByName(name + ' Processing');
  var nextOpen = Number(processing.getRange('D1').getValue());
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
  var found = false;
  
  for (let i=0; i<mlength; i++) {
    if (matches[i][0] == mostRecent[0]) {
      for (let j=0; j < options[3][1]; j++) {
        if (matches[i][2*j+1] == mostRecent[2]) {
          var p1wins = Number(matches[i][2*j+2]) + Number(mostRecent[1]);
          var cell1 = [i+2,9+2*j];
          found = true;
        }
      }
    } else if (matches[i][0] == mostRecent[2]) {
      for (let j=0; j < options[3][1]; j++) {
        if (matches[i][2*j+1] == mostRecent[0]) {
          var p2wins = Number(matches[i][2*j+2]) + Number(mostRecent[3]);
          var cell2 = [i+2,9+2*j];
        }
      }
    }
  }
  if (found && !empty) {
    standings.getRange(cell1[0], cell1[1]+1).setValue(p1wins);
    standings.getRange(cell2[0], cell2[1]+1).setValue(p2wins);
    if (p1wins + p2wins == options[4][1]) {
      standings.getRange(cell1[0], cell1[1]).setFontColor('#38761d');
      standings.getRange(cell2[0], cell2[1]).setFontColor('#38761d');
    } else if ((p1wins + p2wins > options[4][1])) {
      standings.getRange(cell1[0], cell1[1]).setFontColor('#ff0000');
      standings.getRange(cell2[0], cell2[1]).setFontColor('#ff0000');
    } else {
      standings.getRange(cell1[0], cell1[1]).setFontColor('#990000');
      standings.getRange(cell2[0], cell2[1]).setFontColor('#990000');    
    }
    processing.getRange(nextOpen, 1, 2, 3).setValues([[mostRecent[0], mostRecent[1], mostRecent[3]], [mostRecent[2], mostRecent[3], mostRecent[1]]]);
    processing.getRange('D1').setValue(nextOpen + 2);
    sendResultToDiscord(name, 'Standings', standings.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(cell1.toString() + ',' + cell2.toString() + ',' + nextOpen);
  }
  
  return found;
}

function removeRandom(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',').map(x=>Number(x));
  var standings = sheet.getSheetByName(name + ' Standings');
  var processing = sheet.getSheetByName(name + ' Processing');
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  
  var p1wins = Number(standings.getRange(loc[0],loc[1]+1).getValue()) - Number(removal[2]);
  standings.getRange(loc[0],loc[1]+1).setValue(p1wins ? p1wins : '');
  var p2wins = Number(standings.getRange(loc[2],loc[3]+1).getValue()) - Number(removal[4]);
  standings.getRange(loc[2],loc[3]+1).setValue(p2wins ? p2wins : '');
  if (p1wins + p2wins == options[4][1]) {
      standings.getRange(loc[0],loc[1]).setFontColor('#38761d');
      standings.getRange(loc[2],loc[3]).setFontColor('#38761d');
    } else if ((p1wins + p2wins > options[4][1])) {
      standings.getRange(loc[0],loc[1]).setFontColor('#ff0000');
      standings.getRange(loc[2],loc[3]).setFontColor('#ff0000');
    } else {
      standings.getRange(loc[0],loc[1]).setFontColor('#990000');
      standings.getRange(loc[2],loc[3]).setFontColor('#990000');    
    }
  processing.getRange(loc[4], 1, 2, 3).setValues([['','',''], ['','','']]);
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

// Groups 

function groupsOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,7,2).setValues([
    ["''" + name + "' Groups Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Players per Group", ""],
    ["Grouping Method", "pool draw"],
    ["Aggregation Method", "rank then wins"],
    ["Games per Match", ""]
  ]);

  options.getRange(base + 5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['pool draw', 'random', 'snake'], true).build());
  options.getRange(base + 6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['rank then wins', 'separate', 'fixed', 'wins only', 'none'], true).build());
}

function createGroups(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var ppg = Number(options[3][1]);
  var standings = sheet.insertSheet(name + ' Standings');
  standings.deleteColumns(6+ppg, 21-ppg);
  standings.setColumnWidth(1, 150);
  standings.setColumnWidths(2, 4+ppg, 50);
  standings.getRange('E:E').setNumberFormat('0.000');
  var processing = sheet.insertSheet(name + ' Processing');
  var players = Array.from({length:Number(options[2][1])}, (v,i) => "=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (i+2) + ")");
  var ngroups = Math.ceil(options[2][1]/ppg);
  
  //establish groups
  var groups = [];
  var openCounter = 0;
  switch (options[4][1]) {
    case 'random':
      shuffle(players);
      while (players.length % ppg) {players.push('OPEN'+ ++openCounter);}
      break;
    case 'pool draw':
      while (players.length % ppg) {players.push('OPEN'+ ++openCounter);}
      for (let i=0; i<ppg; i++) {
        shuffle(players, ngroups*i, ngroups*(i+1)-1);
      }
    default:
      while (players.length % ppg) {players.push('OPEN'+ ++openCounter);}
  }
  var pllength = players.length;
  processing.getRange('F1').setValue(pllength + 2);
  for (let i=0; i<pllength; i++) {
    if (Math.floor(i/ngroups) % 2) {
      groups.push(/^OPEN\d+$/.test(players[i]) ? [players[i], '', ngroups-1-(i%ngroups), -1, 2] : [players[i], '', ngroups-1-(i%ngroups), 0, 0]);
    } else {
      groups.push(/^OPEN\d+$/.test(players[i]) ? [players[i], '', i%ngroups, -1, 2] : [players[i], '', i%ngroups, 0, 0]); 
    }
  }
  processing.getRange(1, 1, groups.length, 5).setValues(groups);
  processing.getRange('G1').setValue('=query(U:Z, "select U, V, X/' + options[6][1] + ', W, X-W, Y where not U=\'\' order by V, Y desc, Z desc")');
  processing.getRange(2, 13, pllength, 1).setValues(Array.from({length: pllength}, (x,i) => [1+(i%ppg)]));
  processing.getRange('U1').setFormula('=query(A:E, "select A, avg(C), sum(D), sum(D)+sum(E) where not A=\'\' group by A order by avg(C)")');
  processing.getRange('Y2').setFormula('=arrayformula(iferror(W2:W' + (pllength + 1) + '/X2:X' + (pllength + 1) + ',0))');
  var indScores = [];
  var tbs = [];
  for (let i=0; i<pllength; i++) {
    indScores.push([]);
    for (let j=0; j<ppg; j++) {
      indScores[i].push('=iferror(index(query(A:E, "select sum(D) where A=\'"&U'+(i+2)+'&"\' and B=\'"&U'+(i+2+j-(i%ppg))+'&"\'"),2))');
    }
    tbs.push(['=sum(arrayformula(if(Y'+(i+2)+'=transpose(Y'+(2+ppg*Math.floor(i/ppg))+':Y'+(1+ppg*(1+Math.floor(i/ppg)))+'),\
AA'+(i+2)+':'+columnFinder(26+Number(ppg))+(i+2)+',)))']);
  }
  processing.getRange(2, 27, indScores.length, ppg).setValues(indScores);
  processing.getRange(2, 26, tbs.length, 1).setValues(tbs);
  
  //make standings
  var tables = [];
  var groupCounter = 1;
  var procSheet = "'" + name + " Processing'!";
  for (let i=0; i<pllength; i++) {
    if (i % ppg == 0) {
      tables.push(['Group ' + groupCounter++, '', '', '', '', 'Wins Grid'].concat(Array.from({length: ppg-1}, x => '')));
      standings.getRange(tables.length, 1, 1, 5).merge().setHorizontalAlignment('center').setBackground('#bdbdbd');
      standings.getRange(tables.length, 6, 1, ppg).merge().setHorizontalAlignment('center').setBackground('#bdbdbd');
      tables.push(['Player','Played','Wins','Losses','Win %'].concat(Array.from({length: ppg}, (x,j) => '=left(A' + (tables.length + j + 2) + ',3)')));
    }
    if (i % ppg == ppg-1) {
      tables.push(["=IF(regexmatch(" + procSheet + "G" + (i+2) + ",\"^OPEN\\d+$\"),, " + procSheet + "G" + (i+2)+ ")",
        "=IF(regexmatch(" + procSheet + "G" + (i+2) + ",\"^OPEN\\d+$\"),, " + procSheet + "I" + (i+2)+ ")",
        "=IF(regexmatch(" + procSheet + "G" + (i+2) + ",\"^OPEN\\d+$\"),, " + procSheet + "J" + (i+2)+ ")",
        "=IF(regexmatch(" + procSheet + "G" + (i+2) + ",\"^OPEN\\d+$\"),, " + procSheet + "K" + (i+2)+ ")",
        "=IF(regexmatch(" + procSheet + "G" + (i+2) + ",\"^OPEN\\d+$\"),, " + procSheet + "L" + (i+2)+ ")",
      ]);
    } else {
      tables.push(["=" + procSheet + "G" + (i+2), "=" + procSheet + "I" + (i+2), "=" + procSheet + "J" + (i+2), "=" + procSheet + "K" + (i+2), "=" + procSheet + "L" + (i+2)]);
    }
    let last = tables.length - 1;
    for (let j=0; j<ppg; j++) {
      let toprow = 2+ppg*Math.floor(i/ppg);
      let bottomrow = 1+ppg*(1+Math.floor(i/ppg));
      tables[last].push("=iferror(index(" + procSheet + "AA" + toprow + ":" + columnFinder(26+ppg) + bottomrow + 
        ",match(A" + (last+1) + "," + procSheet + "U" + toprow + ":U" + bottomrow +
        ",0),match(A" + (last + 1 + j - (i%ppg)) + "," + procSheet + "U" + toprow + ":U" + bottomrow + ",0)))"
      );
    }
    if (i % ppg == ppg-1) {
      tables.push(Array.from({length: 5+ppg}, x => ''));
      tables.push(Array.from({length: 5+ppg}, x => ''));
    }
  }
  standings.getRange(1, 1, tables.length, 5+ppg).setValues(tables);
  standings.deleteRows(tables.length, 1001-tables.length)
  
  //make aggregates
  function aggFormat() {
    aggregate.setFrozenRows(1);
    aggregate.deleteColumns(9, 18);
    aggregate.setColumnWidth(1, 50);
    aggregate.setColumnWidth(2, 150);
    aggregate.setColumnWidth(3, 50);
    aggregate.setColumnWidth(4, 150);
    aggregate.setColumnWidths(5, 4, 50);
    aggregate.getRange('H:H').setNumberFormat('0.000');
  }  
  switch (options[5][1]) {
    case 'wins only':
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by L desc")');
      var aggregate = sheet.insertSheet(name + ' Aggregation');      
      aggregate.deleteRows(pllength + 2, 999-pllength);
      aggFormat();
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<Number(options[2][1]); i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      break;
    case 'separate':
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, L desc")');
      for (let i=0; i<options[3][1]; i++) {
        var aggregate = sheet.insertSheet(name + ' Aggregation ' + (i+1) + ((i==0) ? 'sts' : (i==1) ? 'nds' : (i==2) ? 'rds' : 'ths'));
        aggregate.deleteRows(ngroups + 2, 999-ngroups);
        aggFormat();
        let aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
        for (let j=0; j<ngroups; j++) {
          let prIn = i*ngroups+j+2;
          if (i == options[3][1]-1) {
            let openCheckString = "=IF(regexmatch(" + procSheet + "N" + prIn + ",\"^OPEN\\d+$\"),,";
            aggTable.push([openCheckString + (j+1) + ")", openCheckString + procSheet + "N" + prIn + ")", openCheckString + procSheet + "O" + prIn + ")",
              openCheckString + "index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + prIn + "))",
              openCheckString + procSheet + "Q" + prIn + ")", openCheckString + procSheet + "R" + prIn + ")", openCheckString + procSheet + "S" + prIn + ")", openCheckString + procSheet + "T" + prIn + ")"
            ]);
          } else {
            aggTable.push([j+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
              "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
              "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
            ]);
          } 
        }
        aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      }
      break;
    case 'fixed':
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, H")');
      var aggregate = sheet.insertSheet(name + ' Aggregation');
      aggFormat();
      aggregate.deleteRows(2+2*ngroups, 999-2*ngroups);
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<ngroups; i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      for(let i=ngroups; i>0; i-=2) {
        let prIn = ngroups+i;
        aggTable.push([2*ngroups+1-i, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
        prIn = ngroups+i+1;
        aggTable.push([2*ngroups+2-i, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      break;
    case 'none':
      break;
    default:
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, L desc")');
      var aggregate = sheet.insertSheet(name + ' Aggregation');
      aggregate.deleteRows(pllength + 2, 999-pllength);
      aggFormat();
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<Number(options[2][1]); i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
  }
}

function updateGroups(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var processing = sheet.getSheetByName(name + ' Processing');
  var nextOpen = Number(processing.getRange('F1').getValue());
  var groups = processing.getRange(1, 1, options[3][1]*Math.ceil(options[2][1]/options[3][1]), 3).getValues();
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
  
  if (mostRecent[0] == mostRecent[2]) {return false;}
  
  for (let i=0; i<options[2][1]; i++) {
    if (groups[i][0] == mostRecent[0]) {
      var p1Group = groups[i][2];
    } else if (groups[i][0] == mostRecent[2]) {
      var p2Group = groups[i][2];
    }
  }
  if (p1Group == p2Group) {
    if (!empty) {
      processing.getRange(nextOpen, 1, 2, 5).setValues([[mostRecent[0], mostRecent[2], p1Group, mostRecent[1], mostRecent[3]], [mostRecent[2], mostRecent[0], p1Group, mostRecent[3], mostRecent[1]]])
      processing.getRange('F1').setValue(nextOpen+2);
      sendResultToDiscord(name, 'Standings',sheet.getSheetByName(name + ' Standings').getSheetId());
      var results = sheet.getSheetByName('Results');
      var resultsLength = results.getDataRange().getNumRows();
      results.getRange(resultsLength, 7).setValue(name);
      results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(nextOpen);
    }
    return true;
  } else {
    return false;
  }
}

function removeGroups(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var processing = sheet.getSheetByName(name + ' Processing');
  
  processing.getRange(removal[7], 1, 2, 5).setValues([['','','','',''],['','','','','']]);
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

// Barrage

function barrageOptions(num, name) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');
  const base = 10*num;

  options.getRange(base + 1,1,5,2).setValues([
    ["''" + name + "' Barrage Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Grouping Method", "pool draw"],
    ["Games to Win", ""]
  ]);

  options.getRange(base + 4,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['pool draw', 'random', 'snake'], true).build());
}

function createBarrage(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1])
  var players = Array.from({length:nplayers}, (v,i) => "=index(indirect(Options!B" + (2+10*num) + "&\"!B:B\")," + (i+2) + ")");

  if (nplayers % 4) {
    let ui = SpreadsheetApp.getUi();
    let response = ui.alert("Warning: the number of players specified for the " + name + " stage is not a multiple of 4", "Barrage stages work bset with multiples of 4. Would you still like to proceed?", ui.ButtonSet.YES_NO);
    if (response != ui.Button.YES) {
      return;
    }
    for (let i = Number(options[2][1]) % 4; i<4; i++) {
      players.push("BYE");
    }
  }

  var ngroups = players.length/4;
  var matches = sheet.insertSheet(name + " Matches");
  matches.deleteColumns(16, 11);
  matches.deleteRows(10*ngroups, 1001-10*ngroups);
  matches.getRange(matches.getMaxRows(), matches.getMaxColumns()).setFontColor("#FFFFFF").setValue("~");

  for (let j=1; j<=12; j++) {
    switch(j%4) {
      case 1:
        matches.setColumnWidth(j, 50);
        break;
      case 2:
        matches.setColumnWidth(j, 150);
        break;
      case 3:
        matches.setColumnWidth(j, 30);
        break;
    }
  }
  matches.setColumnWidth(13, 20);
  matches.setColumnWidth(14, 150);

  switch(options[3][1]) {
    case 'random':
      shuffle(players, 0, nplayers-1);
      break;
    case 'pool draw':
      for (let i=0; i<3; i++) {
        shuffle(players, ngroups*i, ngroups*(i+1)-1);
      }
      shuffle(players, ngroups*3, nplayers-1);
      break;
  }

  for (let i=0; i<ngroups; i++) {
    let base = 10*i;
    //frames
    if (i != 0) {
      matches.getRange("A" + (base) + ":" + (base)).setBackground("#707070");
    }
    matches.getRange(2+base, 2).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue("Match 1");
    matches.getRange(6+base, 2).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue("Match 2");
    matches.getRange(2+base, 6).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue("Match W");
    matches.getRange(6+base, 6).setFontColor('#b7b7b7').setHorizontalAlignment('right').setValue("Match L");
    matches.getRange(3+base, 1, 2, 1).setHorizontalAlignment('right').setValues([[1],[3]]);
    matches.getRange(7+base, 1, 2, 1).setHorizontalAlignment('right').setValues([[2],[4]]);
    matches.getRange(3+base, 5, 2, 1).setHorizontalAlignment('right').setValues([["W 1"],["W 2"]]);
    matches.getRange(7+base, 5, 2, 1).setHorizontalAlignment('right').setValues([["L 1"],["L 2"]]);
    matches.getRange(5+base, 9, 2, 1).setHorizontalAlignment('right').setValues([["L W"],["W L"]]);
    matches.getRange(3+base, 13, 5, 2).setValues([['Pl', 'Player'], [1, ''], [2, ''], [3, ''], [4, '']]);
    matches.getRange(3+base, 2, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    matches.getRange(7+base, 2, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    matches.getRange(3+base, 6, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    matches.getRange(7+base, 6, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    matches.getRange(5+base, 10, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //player placement
    if (players[4*ngroups - i - 1] == "BYE") {
      matches.getRange(3+base, 1, 2, 2).setValues([[1, players[i]], ['', "BYE"]]);
      matches.getRange(3+base, 2, 1, 2).setBackground('#D9EBD3');
      matches.getRange(3+base, 6).setValue(players[i]);
      matches.getRange(7+base, 6).setValue("BYE");
      matches.getRange(7+base, 13).setValue('');
      matches.getRange(7+base, 1, 2, 2).setValues([[2, players[2*ngroups - i - 1]], [3, players[2*ngroups + i]]]);
    } else {
      matches.getRange(3+base, 2, 2, 1).setValues([[players[i]], [players[2*ngroups + i]]]);
      matches.getRange(7+base, 2, 2, 1).setValues([[players[2*ngroups - i - 1]], [players[4*ngroups - i - 1]]]);
    }
  }

  var places = sheet.insertSheet(name + " Places");
  places.deleteColumns(3,24);
  places.setColumnWidth(1,40);
  places.setColumnWidth(2,150);
  places.getRange('A:A').setVerticalAlignment('top').setHorizontalAlignment('right');
  places.getRange(1,1,1,2).setValues([['Place', 'Player']]);
  for (let i=0; i<4; i++) {
    places.getRange(2+i*ngroups, 1).setValue(i+1);
    places.getRange(2+i*ngroups, 1, ngroups, 1).merge();
  }
  places.deleteRows(nplayers + 2, 999-nplayers);
}

function updateBarrage(num, name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[3][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var ngroups = Math.ceil(Number(options[2][1])/4);
  var byecount = 4 - Number(options[2][1]) % 4;
  var matches = sheet.getSheetByName(name + " Matches");
  var places = sheet.getSheetByName(name + " Places");
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
  var found = false;

  for (let i=0; i<ngroups; i++) {
    var base = 10*i;
    var match = [
      matches.getRange(3+base, 2, 2, 2).getValues(),
      matches.getRange(7+base, 2, 2, 2).getValues(),
      matches.getRange(3+base, 6, 2, 2).getValues(),
      matches.getRange(7+base, 6, 2, 2).getValues(),
      matches.getRange(5+base, 10, 2, 2).getValues(),
    ];
    for (let j=0; j<5; j++) {
      switch (j) {
        case 0:
          if (match[2][0][0]) {continue;}
          break;
        case 1:
          if (match[2][1][0]) {continue;}
          break;
        case 2:
          if (match[4][0][0]) {continue;}
          break;
        case 3:
          if (match[4][1][0]) {continue;}
          break;
      }
      if ((match[j][0][0] == mostRecent[0]) && (match[j][1][0] == mostRecent[2])) {
        var topwins = Number(match[j][0][1]) + Number(mostRecent[1]);
        var bottomwins = Number(match[j][1][1]) + Number(mostRecent[3]);
        found = true;
      } else if ((match[j][0][0] == mostRecent[2]) && (match[j][1][0] == mostRecent[0])) {
        var topwins = Number(match[j][0][1]) + Number(mostRecent[3]);
        var bottomwins = Number(match[j][1][1]) + Number(mostRecent[1]);
        found = true;
      }
      if (found) {
        var loc = i + "," + j;
        switch(j) {
          case 0:
            matches.getRange(3+base, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            if ((topwins >= Number(options[4][1])) && (topwins > bottomwins)) {
              matches.getRange(3+base, 2, 1, 2).setBackground('#D9EBD3');
              matches.getRange(3+base, 6).setValue(match[j][0][0]);
              matches.getRange(7+base, 6).setValue(match[j][1][0]);
            } else if ((bottomwins >= Number(options[4][1])) && (bottomwins > topwins)) {
              matches.getRange(4+base, 2, 1, 2).setBackground('#D9EBD3');
              matches.getRange(3+base, 6).setValue(match[j][1][0]);
              matches.getRange(7+base, 6).setValue(match[j][0][0]);
            }
            break;
          case 1:
            matches.getRange(7+base, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            if ((topwins >= Number(options[4][1])) && (topwins > bottomwins)) {
              matches.getRange(7+base, 2, 1, 2).setBackground('#D9EBD3');
              matches.getRange(4+base, 6).setValue(match[j][0][0]);
              matches.getRange(8+base, 6).setValue(match[j][1][0]);
              if (match[3][0][0] == "BYE") {
                matches.getRange(8+base, 6, 1, 2).setBackground('#D9EBD3');
                matches.getRange(6+base, 10).setValue(match[j][1][0]);
              }
            } else if ((bottomwins >= Number(options[4][1])) && (bottomwins > topwins)) {
              matches.getRange(8+base, 2, 1, 2).setBackground('#D9EBD3');
              matches.getRange(4+base, 6).setValue(match[j][1][0]);
              matches.getRange(8+base, 6).setValue(match[j][0][0]);
              if (match[3][0][0] == "BYE") {
                matches.getRange(8+base, 6, 1, 2).setBackground('#D9EBD3');
                matches.getRange(6+base, 10).setValue(match[j][0][0]);
              }
            }
            break;
          case 2:
            matches.getRange(3+base, 7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            if ((topwins >= Number(options[4][1])) && (topwins > bottomwins)) {
              matches.getRange(3+base, 6, 1, 2).setBackground('#D9EBD3');
              matches.getRange(4+base, 14).setValue(match[j][0][0]);
              places.getRange(2+i, 2).setValue(match[j][0][0]);
              matches.getRange(5+base, 10).setValue(match[j][1][0]);
            } else if ((bottomwins >= Number(options[4][1])) && (bottomwins > topwins)) {
              matches.getRange(4+base, 6, 1, 2).setBackground('#D9EBD3');
              matches.getRange(4+base, 14).setValue(match[j][1][0]);
              places.getRange(2+i, 2).setValue(match[j][1][0]);
              matches.getRange(5+base, 10).setValue(match[j][0][0]);
            }
            break;
          case 3:
            matches.getRange(7+base, 7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            if ((topwins >= Number(options[4][1])) && (topwins > bottomwins)) {
              matches.getRange(7+base, 6, 1, 2).setBackground('#D9EBD3');
              matches.getRange(6+base, 10).setValue(match[j][0][0]);
              matches.getRange(7+base, 14).setValue(match[j][1][0]);
              places.getRange(2+3*ngroups+i-byecount, 2).setValue(match[j][1][0]);
            } else if ((bottomwins >= Number(options[4][1])) && (bottomwins > topwins)) {
              matches.getRange(8+base, 6, 1, 2).setBackground('#D9EBD3');
              matches.getRange(6+base, 10).setValue(match[j][1][0]);
              matches.getRange(7+base, 14).setValue(match[j][0][0]);
              places.getRange(2+3*ngroups+i-byecount, 2).setValue(match[j][0][0]);
            }
            break;
          case 4:
            matches.getRange(5+base, 11, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
            if ((topwins >= Number(options[4][1])) && (topwins > bottomwins)) {
              matches.getRange(5+base, 10, 1, 2).setBackground('#D9EBD3');
              matches.getRange(5+base, 14).setValue(match[j][0][0]);
              places.getRange(2+ngroups+i, 2).setValue(match[j][0][0]);
              matches.getRange(6+base, 14).setValue(match[j][1][0]);
              places.getRange(2+2*ngroups+i, 2).setValue(match[j][1][0]);
            } else if ((bottomwins >= Number(options[4][1])) && (bottomwins > topwins)) {
              matches.getRange(6+base, 10, 1, 2).setBackground('#D9EBD3');
              matches.getRange(5+base, 14).setValue(match[j][1][0]);
              places.getRange(2+ngroups+i, 2).setValue(match[j][1][0]);
              matches.getRange(6+base, 14).setValue(match[j][0][0]);
              places.getRange(2+2*ngroups+i, 2).setValue(match[j][0][0]);
            }
            break;
        }
        break;
      }
    }
    if (found) {break;}
  }

  if (found && !empty) {
    sendResultToDiscord(name, 'Results', matches.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(loc);
  }

  return found;
}

function removeBarrage(num, name, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',').map(x => Number(x));
  var options = sheet.getSheetByName('Options').getRange(num*10 + 1,1,10,2).getValues();
  var ngroups = Math.ceil(Number(options[2][1])/4);
  var byecount = 4 - Number(options[2][1]) % 4;
  var matches = sheet.getSheetByName(name + " Matches");
  var places = sheet.getSheetByName(name + " Places");
  var base = 10*loc[0];
  var match = null;

  switch(loc[1]) {
    case 0:
      if (matches.getRange(3+base, 7).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the winner plays in has a result");
        return false;
      } else if (matches.getRange(7+base, 7).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the loser plays in has a result");
        return false;
      }
      matches.getRange(3+base, 6).setValue('');
      matches.getRange(7+base, 6).setValue('');
      match = matches.getRange(3+base, 2, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      matches.getRange(3+base, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      break;
    case 1:
      if (matches.getRange(4+base, 7).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the winner plays in has a result");
        return false;
      } else if (matches.getRange(8+base, 7).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the loser plays in has a result");
        return false;
      } else if ((matches.getRange(7+base, 6).getValue() === "BYE") && (matches.getRange(6+base, 11).getValue() !== '')) {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the loser plays in has a result");
        return false;
      }
      matches.getRange(4+base, 6).setValue('');
      matches.getRange(8+base, 6).setValue('');
      matches.getRange(6+base, 10).setValue('');
      matches.getRange(7+base, 6, 2, 2).setBackground(null);
      match = matches.getRange(7+base, 2, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      matches.getRange(7+base, 3, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      break;
    case 2:
      if (matches.getRange(5+base, 11).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the loser plays in has a result");
        return false;
      }
      matches.getRange(4+base, 14).setValue('');
      places.getRange(2+loc[0], 2).setValue('');
      matches.getRange(5+base, 10).setValue('');
      match = matches.getRange(3+base, 6, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      matches.getRange(3+base, 7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      break;
    case 3:
      if (matches.getRange(6+base, 11).getValue() !== '') {
        SpreadsheetApp.getUi().alert("Removal Failue (Barrage): The match the winner plays in has a result");
        return false;
      }
      matches.getRange(6+base, 10).setValue('');
      matches.getRange(7+base, 14).setValue('');
      places.getRange(2+3*ngroups+loc[0]-byecount, 2).setValue('');
      match = matches.getRange(7+base, 6, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      matches.getRange(7+base, 7, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      break;
    case 4:
      matches.getRange(5+base, 14).setValue('');
      places.getRange(2+ngroups+loc[0], 2).setValue('');
      matches.getRange(6+base, 14).setValue('');
      places.getRange(2+2*ngroups+loc[0], 2).setValue('');
      match = matches.getRange(5+base, 10, 2, 2).setBackground(null).getValues();
      if (removal[1] == match[0][0]) {
        var topwins = Number(match[0][1]) - Number(removal[2]);
        var bottomwins = Number(match[1][1]) - Number(removal[4]);
      } else {
        var topwins = Number(match[0][1]) - Number(removal[4]);
        var bottomwins = Number(match[1][1]) - Number(removal[2]);
      }
      matches.getRange(5+base, 11, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
      break;
  }
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}
