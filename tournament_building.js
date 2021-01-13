// A name and type for each stage, ie [{name: 'Swiss', type: 'swiss'}, {name: 'Elimination', type: 'single'}]
var stages = [];

// The sheet that contains the player list in its second column
var playerListSheet = '';

//If wanting to remove a result
var rowToRemove = 0;

function setupInitial() { 
  adminSetup();
  renameResults();
  addStages();
}

function setupSheets() {
  for (let i=0; i<stages.length; i++) {
    switch (stages[i].type.toLowerCase()) {
      case 'single':
      case 'single elimination':
        createSingleBracket(i);
        break;
      case 'double':
      case 'double elimination':
        createDoubleBracket(i);
        break;
      case 'swiss':
      case 'swiss system':
        createSwiss(i);
        break;
      case 'random':
      case 'preset random':
        createRandom(i);
        break;
      case 'group':
      case 'groups':
        createGroups(i);
        break;
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').setActiveSelection('C7');
}

function removeResult() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var removal = sheet.getSheetByName('Results').getRange(rowToRemove,1,1,8).getValues()[0];
  var success = false;

  var removalStage = 'none';
  for (let i=0; i<stages.length; i++) {
    if (stages[i].name == removal[6]) {
      removalStage = i;
      break;
    }
  }
  if (removalStage != 'none') {
    if (admin[removalStage + 6][2] && !admin[removalStage + 6][3]) {
      switch (stages[removalStage].type.toLowerCase()) {
        case 'single':
        case 'single elimination':
          success = removeSingleBracket(removalStage, rowToRemove);
          break;
        case 'double':
        case 'double elimination':
          success = removeDoubleBracket(removalStage, rowToRemove);
          break;
        case 'swiss':
        case 'swiss system':
          success = removeSwiss(removalStage, rowToRemove);
          break;
        case 'random':
        case 'preset random':
          success = removeRandom(removalStage, rowToRemove);
          break;
        case 'group':
        case 'groups':
          success = removeGroups(removalStage, rowToRemove);
          break;
      }
      if (success) {
        var form = FormApp.openByUrl(admin[2][1].replace(/viewform$/, 'edit'));
        var items = form.getItems();
        var zeroMatch = form.createResponse();
        zeroMatch.withItemResponse(items[0].asListItem().createResponse(removal[1])).withItemResponse(items[1].asTextItem().createResponse('0'));
        zeroMatch.withItemResponse(items[2].asListItem().createResponse(removal[3])).withItemResponse(items[3].asTextItem().createResponse('0'));
        zeroMatch.withItemResponse(items[4].asTextItem().createResponse('Removal ' + rowToRemove + ' - ' + removal[6] + ': ' + removal[1] + ' ' + removal[2] + '-' + removal[4] + ' ' + removal[3]));
        zeroMatch.submit();
      }
    } else {
      Logger.log('Removal Failure: the result is from an inactive stage');
    }
  } else {
    Logger.log('Removal Failure: the result is not in any stage');
  }
}

function testWebhook() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];

  if (webhook) {
    var possibles = ["[Hi! I'm your talking webhook. Do you have a result?](https://www.youtube.com/watch?v=9-QWW2Cxn98)",
                    "[Ever wonder how your results are processed?](https://www.youtube.com/watch?v=OZT9IvH5mC8)"];
    
    var content = possibles[Math.floor(Math.random() * possibles.length)] + "\n\nThis webhook will display " + admin[0][1] + " results as they come in.";
    
    var testMessage = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": null,
        "embeds": [
          {
            "title": "Test and Introduction",
            "color": parseInt('0x'.concat(admin[3][3].slice(1))),
            "description": content
          }
        ]
      })
    };
    
    UrlFetchApp.fetch(webhook, testMessage);
  } else {
    Logger.log("No webhook URL provided")
  }
  
}

function remakeStages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  
  for (let i=0; i<stages.length; i++) {
    if (!admin[i+6][2]) {
      switch (stages[i].type.toLowerCase()) {
        case 'single':
        case 'single elimination':
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Bracket'));
          if (sheet.getSheetByName(stages[i].name + ' Places')) {sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Places'));}
          createSingleBracket(i);
          break;
        case 'double':
        case 'double elimination':
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Bracket'));
          if (sheet.getSheetByName(stages[i].name + ' Places')) {sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Places'));}
          createDoubleBracket(i);
          break;
        case 'swiss':
        case 'swiss system':
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Standings'));
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Processing'));
          createSwiss(i);
          break;
        case 'random':
        case 'preset random':
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Standings'));
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Processing'));
          createRandom(i);
          break;
        case 'group':
        case 'groups':
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Standings'));
          sheet.deleteSheet(sheet.getSheetByName(stages[i].name + ' Processing'));
          var sheets = sheet.getSheets();
          let aggRegex = new RegExp('^' + stages[i].name + ' Aggregation');
          for (let j=sheets.length-1; j>=0; j--) {
            if(aggRegex.test(sheets[j].getSheetName())) {
              sheet.deleteSheet(sheets[j]);
              sheets.splice(j,1);
            }
          }
          createGroups(i);
          break;
      }
    }
  }
  sheet.setActiveSheet(sheet.getSheetByName('Administration'));
}


function adminSetup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  //create a results form, link it, and make a trigger for results submissions
  var form = FormApp.create(sheet.getName() + ' Results Form');
  form.setDescription('Report results here - if you have previously played a partial match, please submit only new games');
  var playerSheet = sheet.getSheetByName(playerListSheet).getDataRange().getValues();
  var names = [];
  for (let i=1; i<playerSheet.length; i++) {
    names.push(playerSheet[i][1]);
  }
  names.sort(function(a,b) {
    if (a.toLowerCase() < b.toLowerCase()) {
      return -1;
    } else if (a.toLowerCase() > b.toLowerCase()) {
      return 1;
    } else if (a.toLowerCase() == b.toLowerCase()) {
      return 0
    }  
  });
  var validation = FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build();
  form.addListItem().setTitle('Your username').setChoiceValues(names).setRequired(true);
  form.addTextItem().setTitle('Your wins').setValidation(validation).setRequired(true);
  form.addListItem().setTitle('Opponent username').setChoiceValues(names).setRequired(true);
  form.addTextItem().setTitle('Opponent wins').setValidation(validation).setRequired(true);
  form.addTextItem().setTitle('Comments');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();
  var formURL = 'https://docs.google.com/forms/d/' + form.getId() + '/viewform';
  
  //add an admin sheet
  var admin = sheet.insertSheet();
  admin.setName('Administration');
  admin.deleteColumns(5, 22);
  admin.deleteRows(7, 994);
  admin.setColumnWidth(1, 150);
  let info = [['Tournament Name', sheet.getName(), '', ''],
              ['Published Spreadsheet', '', '', ''],
              ['Results Form', formURL, '', ''],
              ['Discord Webhook', '', 'Color Hex', ''],
              ['','','',''],
              ['Stage', 'Type', 'Started','Finished']];
  admin.getRange(1, 1, 6, 4).setValues(info);
}

function renameResults() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let i=0; i < sheets.length; i++) {
    if (/^Form Responses/.test(sheets[i].getName())) {
      sheets[i].setName('Results');
      sheets[i].getRange(1, 7).setValue('Stage');
    }
  }
}

function addStages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration');
  var options = sheet.insertSheet('Options');
  options.deleteRows(2,999);
  options.deleteColumns(3, 24);
  options.setColumnWidth(1, 150);
  for (let i=stages.length - 1; i >= 0; i--) {
    admin.insertRowAfter(6);
    admin.getRange('A7:B7').setValues([[stages[i].name, stages[i].type]]);
    options.insertRows(1, 10);
    switch (stages[i].type.toLowerCase()) {
      case 'single':
      case 'single elimination':
        singleBracketOptions(i);
        break;
      case 'double':
      case 'double elimination':
        doubleBracketOptions(i);
        break;
      case 'swiss':
      case 'swiss system':
        swissOptions(i);
        break;
      case 'random':
      case 'preset random':
        var opts = randomOptions(i);
        break;
      case 'group':
      case 'groups':
        var opts = groupsOptions(i);
        break;
    }
    
    options.getRange(2, 1, 9, 2).setBackground(null);
    options.getRange(1, 1, 1, 2).setBackground('#D0E0E3').merge();
  }
  
  admin.getRange(7, 3, stages.length, 2).insertCheckboxes();
  sheet.setActiveSheet(options);
}

function onFormSubmit() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  var found = false;
  
  for (let i=stages.length-1; i>=0; i--) {
    if ((admin[i+6][2]) && (!admin[i+6][3])) {
      switch (stages[i].type.toLowerCase()) {
        case 'single':
        case 'single elimination':
          found = updateSingleBracket(i);
          break;
        case 'double':
        case 'double elimination':
          found = updateDoubleBracket(i);
          break;
        case 'swiss':
        case 'swiss system':
          found = updateSwiss(i);
          break;
        case 'random':
        case 'preset random':
          found = updateRandom(i);
          break;
        case 'group':
        case 'groups':
          found = updateGroups(i);
          break;
      }
    }
    if (found) {break;}
  }
  
  if (!found) {
    sendResultToDiscord('unfound', 0);
    Utilities.sleep(2000);
    var results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();    
    results.getRange(resultsLength, 8).setValue('not found');
  }
}

function sendResultToDiscord(displayType, gid) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];

  var form = FormApp.openByUrl(admin[2][1]).getResponses();
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
            "color": parseInt('0x'.concat(admin[3][3].slice(1))),
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
    display += '\n**[Current ' + displayType + '](' + admin[1][1] + '?gid=' + gid + ')**';
    
    var scoreEmbed = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
      },
      "payload": JSON.stringify({
        "content": null,
        "embeds": [{
          "title": "New " + admin[0][1] + " Result",
          "color": parseInt('0x'.concat(admin[3][3].slice(1))),
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


//Single Elimination functions

function singleBracketOptions(stage) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');

  options.getRange(1,1,8,2).setValues([
    ["''" + stages[stage].name + "' Single Elimination Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Games to Win", ""],
    ["Seeding Method", "standard"],
    ["Consolation", "none"],
    ["Places Sheet", ""],
    ["Winner Count", "1"]
  ]);

  options.getRange(5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'tennis'],true).build());
  options.getRange(6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['none', 'third'],true).build());
  options.getRange(7,2).insertCheckboxes();
}

function createSingleBracket(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[7][1]);
  var nround = Math.ceil(Math.log(nplayers/nbrackets)/Math.log(2));
  var bracketSize = 2**(nround + 1);
  var bracket = sheet.insertSheet(stages[stage].name + ' Bracket');
  bracket.deleteRows(2, 999);
  if ((nbrackets == 1) && (options[5][1] == 'third') && (nround < 4)) {
    bracket.insertRows(1, 3*bracketSize/4 + 5)
  } else {
    bracket.insertRows(1, nbrackets*bracketSize - 1);
  }
  bracket.deleteColumns(2, 25);
  bracket.insertColumns(1, 2*nround + 1);
  
  //make borders
  for (let b=0; b < nbrackets; b++) {
    for (let j=1; j <= nround; j++) {
      let nmatches = 2**(nround-j);
      let space = 2**j;
      for (let i=0; i < nmatches; i++) {
        bracket.getRange(b*bracketSize + space + 2*i*space, 2*j, 2, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        if (j != 1) {
          bracket.getRange(b*bracketSize + space/2 + 2*i*space, 2*j, space, 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        }
      }
    }
  }
  
  if ((nbrackets == 1) && (options[5][1] == 'third')) {
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
  var allSeeds = Array.from({length: nplayers}, (v, i) => (i+1));
  var seedMethod = options[4][1];
  switch (seedMethod) {
    case 'random':
      shuffle(allSeeds);
      break;
    case 'tennis':
      let lastWithBye = nbrackets * 2**nround - nplayers;
      let splitRound = lastWithBye ? Math.ceil(Math.log(lastWithBye/nbrackets)/Math.log(2)) : -1;
      shuffle(allSeeds, nbrackets * 2**(nround-1));      
      for (let i=nround-1; i>0; i--) {
        if (i == splitRound) {
          shuffle(allSeeds, nbrackets * 2**(i-1), lastWithBye - 1);
          shuffle(allSeeds, lastWithBye, nbrackets * 2**i - 1);
        } else {
          shuffle(allSeeds, nbrackets * 2**(i-1), nbrackets * 2**i - 1);
        }
      }
      break;
  }

  var seedList = Array.from({length: nbrackets}, () => []);
  for (let i=0; i<nplayers; i++) {
    if (Math.floor(i/nbrackets) % 2) {
      seedList[nbrackets - 1 - (i%nbrackets)].push(allSeeds[i]);
    } else {
      seedList[i%nbrackets].push(allSeeds[i]);
    }
  }
  
  function placeSeeds(seed, b, location, roundsLeft) {    
    if (2**nround + 1 - seed <= seedList[b].length) {
      if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location, 1).setValue(seedList[b][seed-1]);}
      bracket.getRange(b*bracketSize + location, 2).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[b][seed-1] + 1) + ")");
      if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location + 1, 1).setValue(seedList[b][2**nround - seed]);}
      bracket.getRange(b*bracketSize + location + 1, 2).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[b][2**nround - seed] + 1) + ")");
    } else {
      bracket.getRange(b*bracketSize + location, 2, 2, 2).setBorder(false, false, false, false, null, null);
      if ((location - 2) % 8) {
        if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location - 1, 3).setValue(seedList[b][seed-1]);}
        bracket.getRange(b*bracketSize + location - 1, 4).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[b][seed-1] + 1) + ")");        
      } else {
        if (seedMethod != 'random') {bracket.getRange(b*bracketSize + location + 2, 3).setValue(seedList[b][seed-1]);}
        bracket.getRange(b*bracketSize + location + 2, 4).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[b][seed-1] + 1) + ")");        
      }      
    }
    
    for(let i=nround-roundsLeft+1; i < nround; i++) {
      placeSeeds(2**i + 1 - seed, b, location + 2**(nround+1-i) , nround - i);
    }
    
  }
  for (let b=0; b<nbrackets; b++) {placeSeeds(1,b,2,nround);}
  
  //places sheet if desired
  if (options[6][1]) {
    var places = sheet.insertSheet(stages[stage].name + ' Places');
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
    if ((nbrackets == 1) && (options[5][1] == 'third')) {
      places.getRange(4,1,2,1).breakApart().setValues([[3],[4]]);
    }
    places.deleteRows(nplayers + 2, 999-Number(nplayers));
  }
  
}

function updateSingleBracket(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse();});
  var bracket = sheet.getSheetByName(stages[stage].name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[7][1]);
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
        if (bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).getValue() === '') {
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
                bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[0][0]);
                if ((j == nround - 1) && (nbrackets == 1) && (options[5][1] == 'third')) {
                  bracket.getRange(4 + 3*space + ((i==1)?1:0), 2*j+2).setValue(match[1][0]);
                } else if (options[6][1]) {
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
                  let fillLength = nbrackets*nmatches;
                  let tofill = places.getRange(fillLength+2,2,fillLength,1).getValues().flat();
                  for (let k=0; k<fillLength; k++) {
                    if (!tofill[k]) {
                      places.getRange(fillLength+2+k,2).setValue(match[1][0]);
                      break;
                    }
                  }
                }
              } else if (options[6][1]) {
                let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                bracket.getRange(b*bracketSize + space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).setValue(match[1][0]);
                if ((j == nround - 1) && (nbrackets == 1) && (options[5][1] == 'third')) {
                  bracket.getRange(4 + 3*space + ((i==1)?1:0), 2*j+2).setValue(match[0][0]);
                } else if (options[6][1]) {
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
                  let fillLength = nbrackets*nmatches;
                  let tofill = places.getRange(fillLength+2,2,fillLength,1).getValues().flat();
                  for (let k=0; k<fillLength; k++) {
                    if (!tofill[k]) {
                      places.getRange(fillLength+2+k,2).setValue(match[0][0]);
                      break;
                    }
                  }
                }
              } else if (options[6][1]) {
                let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
          } else if ((j == nround) && (nbrackets == 1) && (options[5][1] == 'third')) {
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
                if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(4,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
              } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
                bracket.getRange(5 + 1.5*space, 2*j, 1, 2).setBackground('#D9EBD3');
                if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(4,2,2,1).setValues([[match[1][0]],[match[0][0]]]);}
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
    sendResultToDiscord('Bracket', bracket.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(stages[stage].name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(loc);
  }
  
  return found;
}

function removeSingleBracket(stage, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',');
  var bracket = sheet.getSheetByName(stages[stage].name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nplayers = Number(options[2][1]);
  var nbrackets = Number(options[7][1]);
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
    if (options[6][1]) {
      sheet.getSheetByName(stages[stage].name + ' Places').getRange(4,2,2,1).setValues([[''],['']]);
    }
  } else {
    var space = 2**loc[1];
    
    //scrub dependencies on this match and fail if one has a result
    var toScrub = [];
    if (bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+3).getValue() !== '') {
      Logger.log('Removal Failure (Single Elimination): The match the winner plays in has a result');
      return false;
    } else {
      toScrub.push(bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1] + 2));
    }
    if ((loc[1] == nround-1) && (nbrackets == 1) && (options[5][1] == 'third')) {
      if (bracket.getRange(4 + 3*space, 2*nround+1).getValue() !== '') {
        Logger.log('Removal Failure (Single Elimination): The third place match has a result');
        return false;
      } else {
        toScrub.push(bracket.getRange(4 + 3*space + ((loc[2]==1) ? 1 : 0), 2*loc[1] + 2));
      }
    }
    for (let i=toScrub.length; i>0; i--) {toScrub[i-1].setValue('');}
    
    var match = bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2], 2*loc[1], 2, 2).setBackground(null).getValues();
    if (removal[1] == match[0][0]) {
      var topwins = Number(match[0][1]) - Number(removal[2]);
      var bottomwins = Number(match[1][1]) - Number(removal[4]);
    } else {
      var topwins = Number(match[0][1]) - Number(removal[4]);
      var bottomwins = Number(match[1][1]) - Number(removal[2]);
    }
    bracket.getRange(loc[0]*bracketSize + space + 2*space*loc[2], 2*loc[1] + 1, 2, 1).setValues(((topwins == 0) && (bottomwins == 0)) ? [[''],['']] : [[topwins],[bottomwins]]);
    if (options[6][1]) {
      let places = sheet.getSheetByName(stages[stage].name + ' Places');
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


//Double Elimination functions

function doubleBracketOptions(stage) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');

  options.getRange(1,1,7,2).setValues([
    ["''" + stages[stage].name + "' Double Elimination Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Games to Win", ""],
    ["Seeding Method", "standard"],
    ["Finals", "standard"],
    ["Places Sheet", ""]
  ]);

  options.getRange(5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'tennis'],true).build());
  options.getRange(6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'grand'],true).build());
  options.getRange(7,2).insertCheckboxes();
}

function createDoubleBracket(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var nlround = 2*nround - 2 - (2**nround-options[2][1]>=2**(nround-2));
  var width = Math.max(2*nround + 5, 2*nlround + 1)
  var bracket = sheet.insertSheet(stages[stage].name + ' Bracket');
  bracket.deleteRows(2, 999);
  bracket.insertRows(1, 2**(nround+1) + ((nlround % 2) ? 4 * 2**((nlround-1)/2) : 3 * 2**(nlround/2)));
  bracket.deleteColumns(2, 25);
  bracket.insertColumns(1, width);
  
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
      bracket.getRange(location, 2).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[seed-1] + 1) + ")");
      if (seedMethod != 'random') {bracket.getRange(location + 1, 1).setValue(seedList[2**nround - seed]);}
      bracket.getRange(location + 1, 2).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[2**nround - seed] + 1) + ")");
    } else {
      bracket.getRange(location, 2, 2, 2).setBorder(false, false, false, false, null, null);
      bracket.getRange(location-1,2).setValue('');
      if ((location - 2) % 8) {
        if (seedMethod != 'random') {bracket.getRange(location - 1, 3).setValue(seedList[seed-1]);}
        bracket.getRange(location - 1, 4).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[seed-1] + 1) + ")");        
      } else {
        if (seedMethod != 'random') {bracket.getRange(location + 2, 3).setValue(seedList[seed-1]);}
        bracket.getRange(location + 2, 4).setFormula("=index('" + options[1][1] + "'!B:B," + (seedList[seed-1] + 1) + ")");        
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
    var places = sheet.insertSheet(stages[stage].name + ' Places');
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

function updateDoubleBracket(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var bracket = sheet.getSheetByName(stages[stage].name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
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
        if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(2,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
      } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
        bracket.getRange(2**nround + 3, 2*nround + 4, 1, 2).setBackground('#D9EBD3');
        if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(2,2,2,1).setValues([[match[1][0]],[match[0][0]]]);}
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
          if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(2,2,2,1).setValues([[match[0][0]],[match[1][0]]]);}
        } else if ((bottomwins >= gamesToWin) && (bottomwins > topwins)) {
          bracket.getRange(2**nround + 3, 2*nround + 2, 1, 2).setBackground('#D9EBD3');
          if (options[5][1] != 'grand') {
            bracket.getRange(2**nround + 2, 2*nround + 4, 2, 1).setValues([[match[0][0]], [match[1][0]]]);
          } else if (options[6][1]) {
            sheet.getSheetByName(stages[stage].name + ' Places').getRange(2,2,2,1).setValues([[match[1][0]],[match[0][0]]]);
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
        if (bracket.getRange(space + 2*i*space + ((i%2) ? 1-space : space), 2*j+2).getValue() === '') {
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                    let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                    let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                  let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                    let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
                    let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
    sendResultToDiscord('Bracket', bracket.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(stages[stage].name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(loc);    
  }
  return found;
}

function removeDoubleBracket(stage, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',');
  var bracket = sheet.getSheetByName(stages[stage].name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
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
        Logger.log('Removal Failure (Double Elimination): The second finals match has a result');
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
    if (options[6][1]) {sheet.getSheetByName(stages[stage].name + ' Places').getRange(2,2,2,1).setValues([[''],['']]);}
  } else if (loc[0] == 'w') {
    //in winners' bracket
    var space = 2**loc[1];
    var nmatches = 2**(nround-loc[1]);
    //scrub dependencies on this match and fail if one has a result
    var toScrub = [];
    if (bracket.getRange(space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+3).getValue() !== '') {
      Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
      return;
    } else {
      toScrub.push(bracket.getRange(space + 2*space*loc[2] + ((loc[2]%2) ? 1-space : space),2*loc[1]+2));
    }
    if (nlround % 2) {
      if (loc[1]>2) {
        if (bracket.getRange(ltop + 4 * (2**(loc[1]-3) - 1) - 2*loc[1] + 7 + ((loc[1]%2) ? loc[2] : -1-loc[2]+nmatches)*8*2**(loc[1]-3), 4*loc[1]-5).getValue() !== '') {
          Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 4 * (2**(loc[1]-3) - 1) - 2*loc[1] + 7 + ((loc[1]%2) ? loc[2] : -1-loc[2]+nmatches)*8*2**(loc[1]-3), 4*loc[1]-6));
        }
      } else if (loc[1] == 2) {
        if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,1).getValue()) {
          if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,3).getValue() !== '') {
            Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
            return;
          } else {
            toScrub.push(bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3,2));
          }
        } else {
          if (bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3 + ((loc[2]%2) ? 2 : -1), 5).getValue() !== '') {
            Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
            return;
          } else {
            toScrub.push(bracket.getRange(ltop + 4*(nmatches - loc[2]) - 3 + ((loc[2]%2) ? 2 : -1), 4));
          }
        }
      } else {
        if (bracket.getRange(ltop + 2*loc[2], 3).getValue() !== '') {
          Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 2*loc[2], 2));
        }
      }
    } else {
      if (loc[1]>1) {
        if (bracket.getRange(ltop + 3 * (2**(loc[1]-2) - 1) - 2*loc[1] + 5 + ((loc[2]%2) ? loc[2] : -1-loc[2]+2**(nround-loc[1]))*6*2**(loc[1]-2), 4*loc[1]-3).getValue() !== '') {
          Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
          return;
        } else {
          toScrub.push(bracket.getRange(ltop + 3 * (2**(loc[1]-2) - 1) - 2*loc[1] + 5 + ((loc[2]%2) ? loc[2] : -1-loc[2]+2**(nround-loc[1]))*6*2**(loc[1]-2), 4*loc[1]-4));
        }
      } else {
        if (loc[2]%2) {
          if (bracket.getRange(ltop + 1 + 3*loc[2], 1).getValue()) {
            if (bracket.getRange(ltop + 1 + 3*loc[2], 3).getValue() !== '') {
              Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
              return;
            } else {
              toScrub.push(bracket.getRange(ltop + 1 + 3*loc[2], 2));
            }
          } else {
            if (bracket.getRange(ltop - 1 + 3*loc[2], 5).getValue() !== '') {
              Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
              return;
            } else {
              toScrub.push(bracket.getRange(ltop - 1 + 3*loc[2], 4));
            }
          }
        } else {
          if (bracket.getRange(ltop + 3 + 3*loc[2], 3).getValue() !== '') {
            Logger.log('Removal Failure (Double Elimination): The match the loser plays in has a result');
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
          Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
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
          let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
              Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
              return false;
            } else {
              bracket.getRange(2**nround+3, 2*nround+2).setValue('');
              bracket.getRange(ltop-1, 2).setValue('');
            }
          } else {
            if (bracket.getRange(ltop + topShape + loc[2]*spacing + 2 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+9).getValue() !== '') {
              Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
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
            let places = sheet.getSheetByName(stages[stage].name + ' Places');
            let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues().flat();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+2*nshapes,2).setValue('');
              }
            }
          }
        } else {
          if (bracket.getRange(ltop + topShape + loc[2]*spacing + 3, 4*loc[1]+7).getValue() !== '') {
            Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
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
            let places = sheet.getSheetByName(stages[stage].name + ' Places');
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
            Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
            return false;
          } else {
            bracket.getRange(2**nround+3, 2*nround+2).setValue('');
            bracket.getRange(ltop-1, 2).setValue('');
          }
        } else {
          if (bracket.getRange(ltop + topShape + loc[2]*spacing + 1 + ((loc[2]%2) ? 1 - spacing/2 : spacing/2), 4*loc[1]+7).getValue() !== '') {
            Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
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
            let places = sheet.getSheetByName(stages[stage].name + ' Places');
            let tofill = places.getRange(2+2*nshapes,2,nshapes,1).getValues();
            for (let k=0; k<nshapes; k++) {
              if ((tofill[k] == match[0][0]) || (tofill[k] == match[1][0])) {
                places.getRange(2+k+2*nshapes,2).setValue('');
              }
            }
          }
      } else {
        if (bracket.getRange(ltop + topShape + loc[2]*spacing + 2, 4*loc[1]+5).getValue() !== '') {
          Logger.log('Removal Failure (Double Elimination): The match the winner plays in has a result');
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
            let places = sheet.getSheetByName(stages[stage].name + ' Places');
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


//Swiss system functions

function swissOptions(stage) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');

  options.getRange(1,1,6,2).setValues([
    ["''" + stages[stage].name + "' Swiss Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""],
    ["Seeding Method", "standard"]
  ]);

  options.getRange(6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['standard', 'random', 'highlow'],true).build());
}

function createSwiss(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var swiss = sheet.insertSheet(stages[stage].name + ' Standings');
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
  swiss.getRange('H1').setFontColor('#ffffff');
  swiss.getRange('I:I').setHorizontalAlignment('right');
  swiss.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISBLANK(B1)').setFontColor('#FFFFFF').setRanges([swiss.getRange('A:A')]).build()]);
  var processing = sheet.insertSheet(stages[stage].name + ' Processing');
  processing.getRange('E1').setValue('2');
  var players = [];
  for (let i=0; i<options[2][1]; i++) {
    players.push("=index('" + options[1][1] + "'!B:B," + (i+2) + ")");
  }
  
  //display sheet
  //standings
  var standings = [['Pl', 'Player', 'Wins', 'Losses']];
  for (let i=1; i <= options[2][1]; i++) {
    standings.push([i, "='" + stages[stage].name + " Processing'!G" + (i+1), "='" + stages[stage].name + " Processing'!H" + (i+1), "='" + stages[stage].name + " Processing'!I" + (i+1)]);
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
  var firstround = [['Round', 1, '=SUM(G2:H)/' + (Number(options[4][1]) * halfplayers), '']];
  for (let i = 0; i < halfplayers; i++) {
    if (players[i] == 'BYE') {
      firstround.push(['BYE', 0, options[4][1], players[halfplayers+i]]);
    } else if (players[halfplayers + i] == 'BYE') {
      firstround.push([players[i], options[4][1], 0, 'BYE']);
    } else {
      firstround.push([players[i], '', '', players[halfplayers+i]]);
    }
    
  }
  swiss.getRange(1, 6, firstround.length, 4).setValues(firstround);
  
  //processing sheet
  processing.getRange('G1').setValue('=query(A:D, "select A, sum(C), sum(D) where not A = \'BYE\' group by A order by sum(C) desc")');
  if(byeplace) {
    let playerToInsert = (byeplace < halfplayers) ? players[halfplayers + byeplace] : players[byeplace - halfplayers];
    processing.getRange(2, 1, 1, 4).setValues([[playerToInsert, 'BYE', options[4][1], 0]]);
    processing.getRange(3, 1, 1, 4).setValues([['BYE', playerToInsert, 0, options[4][1]]]);
    processing.getRange('E1').setValue(4);
    processing.getRange('E2').setValue(2);
  }
}

function updateSwiss(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var swiss = sheet.getSheetByName(stages[stage].name + ' Standings');
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  var nextOpen = Number(processing.getRange('E1').getValue());
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  
  //update matches and standings - standings updates only occur if the reported match was in the match list
  var roundMatches = swiss.getRange(2, 6, Math.ceil(options[2][1]/2),4).getValues();
  var empty = (mostRecent[1] == 0) && (mostRecent[3] == 0);
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
      processing.getRange(nextOpen, 1, 2, 4).setValues([[mostRecent[0], mostRecent[2], mostRecent[1], mostRecent[3]], [mostRecent[2], mostRecent[0], mostRecent[3], mostRecent[1]]]);
      processing.getRange('E1').setValue(nextOpen + 2);
      sendResultToDiscord('Standings', swiss.getSheetId());
      var results = sheet.getSheetByName('Results');
      var resultsLength = results.getDataRange().getNumRows();
      results.getRange(resultsLength, 7).setValue(stages[stage].name);
      results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(swiss.getRange('G1').getValue() + ',' + i + ',' + nextOpen);
      //make next round's match list if needed
      newSwissRound(stage);
      break;
    }
  }
  
  return found;
}

function newSwissRound(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var swiss = sheet.getSheetByName(stages[stage].name + ' Standings');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var roundInfo = swiss.getRange('G1:H1').getValues()[0];
  
  if ((roundInfo[1] == 1) && (Number(roundInfo[0]) < Number(options[3][1]))) {
    var players = swiss.getRange(2, 2, options[2][1], 1).getValues().flat();
    var assigned = [];
    var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
    var playedMatches = processing.getRange('A:B').getValues();
    
    //trim playedMatches
    var pmlen = playedMatches.length;
    for (let i=1; i < pmlen; i++) {
      if (!playedMatches[i][0]) {
        playedMatches.splice(i,pmlen-i)
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
    
    var roundMatches = [['Round', Number(roundInfo[0]) + 1, '=SUM(G2:H)/' + (Number(options[4][1]) * halfplayers), '']];
    for (let i=0; i<halfplayers; i++) {
      if (assigned[2*i] == 'BYE') {
        roundMatches.push(['BYE', 0, options[4][1], assigned[2*i+1]]);
      } else if (assigned[2*i+1] == 'BYE') {
        roundMatches.push([assigned[2*i], options[4][1], 0, 'BYE']);
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
    swiss.getRange('H1').setFontColor('#ffffff');
    swiss.getRange('I:I').setHorizontalAlignment('right');
    swiss.getRange('J:J').setBackground('#707070');
    swiss.getRange(1, 6, roundMatches.length, 4).setValues(roundMatches);
    
    if (byeplace) {
      let byeIdx = assigned.indexOf('BYE');
      let nextOpen = Number(processing.getRange('E1').getValue());      
      if (byeIdx % 2) {
        processing.getRange(nextOpen, 1, 2, 4).setValues([[assigned[byeIdx-1],'BYE', options[4][1], 0],['BYE', assigned[byeIdx-1], 0, options[4][1]]]);
      } else {
        processing.getRange(nextOpen, 1, 2, 4).setValues([[assigned[byeIdx+1],'BYE', options[4][1], 0],['BYE', assigned[byeIdx+1], 0, options[4][1]]]);
      }
      processing.getRange('E1').setValue(nextOpen + 2);
      processing.getRange(Number(roundInfo[0]+2), 5).setValue(nextOpen);
    }
  }
}

function removeSwiss(stage, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',').map(x=>Number(x));
  var swiss = sheet.getSheetByName(stages[stage].name + ' Standings');
  var round = Number(swiss.getRange('G1').getValue());
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  
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
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}


//Preset Random functions

function randomOptions(stage) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');

  options.getRange(1,1,5,2).setValues([
    ["''" + stages[stage].name + "' Preset Random Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""]
  ]);
}

function createRandom(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  
  //setup matches
  var standings = sheet.insertSheet(stages[stage].name + ' Standings');
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
  var players = [];
  for (let i=0; i<options[2][1]; i++) {
    players.push("=index('" + options[1][1] + "'!B:B," + (i+2) + ")");
  }
  if (options[2][1] % 2) {players.push('OPEN');}
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
          createRandom(stage);
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
  var processing = sheet.insertSheet(stages[stage].name + ' Processing');
  processing.getRange('D1').setValue(1);
  processing.getRange('E1').setFormula('=query(K:N,"select K,L/M,L,M-L,N where not K = \'BYE\' order by L/M desc, N desc, L desc")');
  processing.getRange('K1').setFormula('=query(A:C,"select A, sum(B), sum(B)+sum(C) group by A order by A")');
  var sosFormulas = [];
  for (let i=0; i<pllength+Boolean(byePlayer); i++) {
    let rowformulas = ["=IFERROR((O"+(i+3)+"+L"+(i+3)+"-M"+(i+3)+")/(P"+(i+3)+"-M"+(i+3)+"),\"NA\")", "=", "="];
    for (let j=0; j<options[3][1]; j++) {
      rowformulas[1] += "+".repeat(j!=0) + "IFERROR(VLOOKUP(VLOOKUP(K" + (i+3) + ",'" + stages[stage].name + " Standings'!H2:" + (options[2][1]+1) + "," + (2*j+2) + "),K:M,2,FALSE),0)";
      rowformulas[2] += "+".repeat(j!=0) + "IFERROR(VLOOKUP(VLOOKUP(K" + (i+3) + ",'" + stages[stage].name + " Standings'!H2:" + (options[2][1]+1) + "," + (2*j+2) + "),K:M,3,FALSE),0)";
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
    table.push([i, "='"+stages[stage].name+" Processing'!E"+(i+1), "='"+stages[stage].name+" Processing'!F"+(i+1),
      "='"+stages[stage].name+" Processing'!G"+(i+1), "='"+stages[stage].name+" Processing'!H"+(i+1), "='"+stages[stage].name+" Processing'!I"+(i+1)]);
  }
  standings.getRange(1,1,table.length,6).setValues(table);
}

function updateRandom(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var standings = sheet.getSheetByName(stages[stage].name + ' Standings');
  var matches = standings.getRange(2, 8, options[2][1], 2*options[3][1]+1).getValues();
  var mlength = matches.length;
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
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
    sendResultToDiscord('Standings', standings.getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(stages[stage].name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(cell1.toString() + ',' + cell2.toString() + ',' + nextOpen);
  }
  
  return found;
}

function removeRandom(stage, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var loc = removal[7].split(',').map(x=>Number(x));
  var standings = sheet.getSheetByName(stages[stage].name + ' Standings');
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  
  var p1wins = Number(standings.getRange(loc[0],loc[1]+1).getValue()) - Number(removal[2]);
  standings.getRange(loc[0],loc[1]+1).setValue(p1wins ? p1wins : '');
  var p2wins = Number(standings.getRange(loc[2],loc[3]+1).getValue()) - Number(removal[4]);
  standings.getRange(loc[2],loc[3]+1).setValue(p2wins ? p2wins : '');
  processing.getRange(loc[4], 1, 2, 3).setValues([['','',''], ['','','']]);
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}


//Groups functions

function groupsOptions(stage) {
  var options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options');

  options.getRange(1,1,7,2).setValues([
    ["''" + stages[stage].name + "' Groups Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Players per Group", ""],
    ["Grouping Method", "pool draw"],
    ["Aggregation Method", "rank then wins"],
    ["Games per Match", ""]
  ]);

  options.getRange(5,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['pool draw', 'random', 'snake'], true).build());
  options.getRange(6,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['rank then wins', 'separate', 'fixed', 'wins only', 'none'], true).build());
}

function createGroups(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var ppg = Number(options[3][1]);
  var standings = sheet.insertSheet(stages[stage].name + ' Standings');
  standings.deleteColumns(6+ppg, 21-ppg);
  standings.setColumnWidth(1, 150);
  standings.setColumnWidths(2, 4+ppg, 50);
  standings.getRange('E:E').setNumberFormat('0.000');
  var processing = sheet.insertSheet(stages[stage].name + ' Processing');
  var players = [];
  for (let i=0; i<options[2][1]; i++) {
    players.push("=index('" + options[1][1] + "'!B:B," + (i+2) + ")");
  }
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
  var procSheet = "'" + stages[stage].name + " Processing'!";
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
        "),match(A" + (last + 1 + j - (i%ppg)) + "," + procSheet + "U" + toprow + ":U" + bottomrow + ")))"
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
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');      
      aggregate.deleteRows(pllength + 2, 999-pllength);
      aggFormat();
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<Number(options[2][1]); i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      break;
    case 'separate':
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, L desc")');
      for (let i=0; i<options[3][1]; i++) {
        var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation ' + (i+1) + ((i==0) ? 'sts' : (i==1) ? 'nds' : (i==2) ? 'rds' : 'ths'));
        aggregate.deleteRows(ngroups + 2, 999-ngroups);
        aggFormat();
        let aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
        for (let j=0; j<ngroups; j++) {
          let prIn = i*ngroups+j+2;
          if (i == options[3][1]-1) {
            let openCheckString = "=IF(regexmatch(" + procSheet + "N" + prIn + ",\"^OPEN\\d+$\"),,";
            aggTable.push([openCheckString + (j+1) + ")", openCheckString + procSheet + "N" + prIn + ")", openCheckString + procSheet + "O" + prIn + ")",
              openCheckString + "index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + prIn + "))",
              openCheckString + procSheet + "Q" + prIn + ")", openCheckString + procSheet + "R" + prIn + ")", openCheckString + procSheet + "S" + prIn + ")", openCheckString + procSheet + "T" + prIn + ")"
            ]);
          } else {
            aggTable.push([j+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
              "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
              "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
            ]);
          } 
        }
        aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      }
      break;
    case 'fixed':
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, H")');
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');
      aggFormat();
      aggregate.deleteRows(2+2*ngroups, 999-2*ngroups);
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<ngroups; i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      for(let i=ngroups; i>0; i-=2) {
        let prIn = ngroups+i;
        aggTable.push([2*ngroups+1-i, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
        prIn = ngroups+i+1;
        aggTable.push([2*ngroups+2-i, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
      break;
    case 'none':
      break;
    default:
      processing.getRange('N1').setFormula('=query(G:M, "select G,M,H,I,J,K,L where not G=\'\' order by M, L desc")');
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');
      aggregate.deleteRows(pllength + 2, 999-pllength);
      aggFormat();
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Played', 'Wins', 'Losses', 'Win %']];
      for (let i=0; i<Number(options[2][1]); i++) {
        let prIn = i+2;
        aggTable.push([i+1, "=" + procSheet + "N" + prIn, "=" + procSheet + "O" + prIn,
          "=index('" + stages[stage].name + " Standings'!A:A, 1+" + (4 + ppg) + "*" + procSheet + "P" + (prIn) + ")",
          "=" + procSheet + "Q" + prIn, "=" + procSheet + "R" + prIn, "=" + procSheet + "S" + prIn, "=" + procSheet + "T" + prIn
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 8).setValues(aggTable);
  }
}

function updateGroups(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
  var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse()});
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
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
  if ((p1Group == p2Group) && !empty) {
    processing.getRange(nextOpen, 1, 2, 5).setValues([[mostRecent[0], mostRecent[2], p1Group, mostRecent[1], mostRecent[3]], [mostRecent[2], mostRecent[0], p1Group, mostRecent[3], mostRecent[1]]])
    processing.getRange('F1').setValue(nextOpen+2);
    sendResultToDiscord('Standings',sheet.getSheetByName(stages[stage].name + ' Standings').getSheetId());
    var results = sheet.getSheetByName('Results');
    var resultsLength = results.getDataRange().getNumRows();
    results.getRange(resultsLength, 7).setValue(stages[stage].name);
    results.getRange(resultsLength, 8).setFontColor('#ffffff').setValue(nextOpen);
    return true;
  } else {
    return false;
  }
}

function removeGroups(stage, row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results');
  var removal = results.getRange(row, 1, 1, 8).getValues()[0];
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  
  processing.getRange(removal[7], 1, 2, 5).setValues([['','','','',''],['','','','','']]);
  results.getRange(row, 7, 1, 2).setValues([['','removed']]).setFontColor(null);
  return true;
}

