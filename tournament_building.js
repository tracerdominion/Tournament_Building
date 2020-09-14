// A name and type for each stage, ie [{name: 'Swiss', type: 'swiss'}, {name: 'Elimination', type: 'single'}]
var stages = [{name: 'Group', type: 'group'}, {name: 'Elimination', type: 'single'}];

// The sheet that contains the player list in its second column
var playerListSheet = 'Player List';

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
}

function removeLastResult() {
  var results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results').getDataRange().getValues();
  var lastResultStage = 'none';
  for (let i=0; i<stages.length; i++) {
    if (stages[i].name == results[results.length-1][6]) {
      lastResultStage = i;
    }
  }
  switch (stages[lastResultStage].type.toLowerCase()) {
    case 'single':
      updateSingleBracket(lastResultStage, false);
      break;
    case 'swiss':
    case 'swiss system':
      updateSwiss(lastResultStage, false);
      break;
    case 'random':
    case 'preset random':
      updateRandom(lastResultStage, false);
      break;
    case 'group':
    case 'groups':
      updateGroups(lastResultStage, false);
      break;
    default:
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results').deleteRow(results.length);
  }
}

function testWebhook() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];
  
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
}

function remakeStages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = sheet.getSheets();
  var nsheets = sheets.length;
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  
  for (let i=0; i<stages.length; i++) {
    if (!admin[i+6][2]) {
      let stageNameRegex = new RegExp('^' + stages[i].name + ' ', 'g');
      for (let j=0; j<nsheets; j++) {
        Logger.log(sheets[j].getSheetName())
        if (stageNameRegex.test(sheets[j].getSheetName())) {
          sheet.deleteSheet(sheets[j]);
        }
      }
      switch (stages[i].type.toLowerCase()) {
        case 'single':
        case 'single elimination':
          createSingleBracket(i);
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
  }
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
  names.sort(lowerCaseOrder);
  var validation = FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build();
  form.addListItem().setTitle('Your username').setChoiceValues(names);
  form.addTextItem().setTitle('Your wins').setValidation(validation);
  form.addListItem().setTitle('Opponent username').setChoiceValues(names);
  form.addTextItem().setTitle('Opponent wins').setValidation(validation);
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
        var opts = singleBracketOptions(i);
        break;
      case 'swiss':
      case 'swiss system':
        var opts = swissOptions(i);
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
    
    options.getRange(1, 1, opts.length, 2).setValues(opts);
    options.getRange(2, 1, 9, 2).setBackground(null);
    options.getRange(1, 1, 1, 2).setBackground('#D0E0E3').merge();
  }
  
  admin.getRange(7, 3, stages.length, 2).insertCheckboxes();
}

function onFormSubmit() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  
  for (let i=0; i < stages.length; i++) {
    if ((admin[i+6][2]) && (!admin[i+6][3])) {
      switch (stages[i].type.toLowerCase()) {
        case 'single':
        case 'single elimination':
          updateSingleBracket(i, true);
          break;
        case 'swiss':
        case 'swiss system':
          updateSwiss(i, true);
          break;
        case 'random':
        case 'preset random':
          updateRandom(i, true);
          break;
        case 'group':
        case 'groups':
          updateGroups(i, true);
          break;
      }
    }
  }
}

function sendResultToDiscord(displayType, gid) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  
  var display = '**' + mostRecent[1] + ' ' + mostRecent[2] + '-' + mostRecent[4] + ' ' + mostRecent[3] + '**';
  if (mostRecent[5]) {
    display += '\n' + mostRecent[5];
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
      "content": "!match " + mostRecent[1] + ", " + mostRecent[3] + " -c"
    })
  };
  
  UrlFetchApp.fetch(webhook, scoreEmbed);
  UrlFetchApp.fetch(webhook, matchCommand);
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

function lowerCaseOrder(a,b) {
  if (a.toLowerCase() < b.toLowerCase()) {
    return -1;
  } else if (a.toLowerCase() > b.toLowerCase()) {
    return 1;
  } else if (a.toLowerCase() == b.toLowerCase()) {
    return 0
  }
}

//Single Elimination functions

function singleBracketOptions(stage) {
  return [
    ["''" + stages[stage].name + "' Single Elimination Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Games to Win", ""],
    ["Seeding Method", "standard"]
  ];
}

function createSingleBracket(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var bracket = sheet.insertSheet(stages[stage].name + ' Bracket');
  bracket.deleteRows(2, 999);
  bracket.insertRows(1, 2**(nround+1) + 2)
  
  //make borders
  for (let j=1; j <= nround + 1; j++) {
    let nplayer = 2**(nround+1-j);
    let offset = 2**(j-1) - 1
    for (let i=1; i <= nplayer; i++) {
      bracket.getRange(i*(2**j) - offset, 2*j, 1, 2).setBorder(true, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      if (j != 1) {
        bracket.getRange(i*(2**j) - offset - 2**(j-2), 2*j, 2**(j-1), 1).setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }   
    }
  }
  
  //column widths
  for (let j=1; j <= nround*2 + 4; j++) {
    if (j % 2) {
      bracket.setColumnWidth(j, 30);
    } else {
      bracket.setColumnWidth(j, 150);
    }
  }
  
  //seed numbers and players
  var seedList = Array.from({length: options[2][1]}, (v, i) => (i+1));
  var seedMethod = options[4][1].toLowerCase();
  switch (seedMethod) {
    case 'random':
      shuffle(seedList);
      Logger.log(seedList);
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
      bracket.getRange(location, 2).setFormula("='" + options[1][1] + "'!B" + (seedList[seed-1] + 1));
      if (seedMethod != 'random') {bracket.getRange(location + 2, 1).setValue(seedList[2**nround - seed]);}
      bracket.getRange(location + 2, 2).setFormula("='" + options[1][1] + "'!B" + (seedList[2**nround - seed] + 1));
    } else {
      if (seedMethod != 'random') {bracket.getRange(location + 1, 3).setValue(seedList[seed-1]);}
      bracket.getRange(location + 1, 4).setFormula("='" + options[1][1] + "'!B" + (seedList[seed-1] + 1));
      bracket.getRange(location, 2, 1, 2).setBorder(false, false, false, false, null, null);
      bracket.getRange(location + 2, 2, 1, 2).setBorder(false, false, false, false, null, null);
    }
    
    for(let i=nround-roundsLeft+1; i < nround; i++) {
      placeSeeds(2**i + 1 - seed, location + 2**(nround+1-i) , nround - i);
    }
    
  }
  placeSeeds(1,2,nround);
}

function updateSingleBracket(stage, add) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  var bracket = sheet.getSheetByName(stages[stage].name + ' Bracket');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var nround = Math.ceil(Math.log(options[2][1])/Math.log(2));
  var gamesToWin = options[3][1];
  var found = false;
  
  for (let j=nround; j>0; j--) {
    let nplayer = 2**(nround+1-j);
    let offset = 2**(j-1) - 1
    for (let i=1; i<=nplayer; i++) {
      let lower = bracket.getRange(i*(2**(j+1)) - offset, 2*j).getValue();
      let upper = bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j).getValue();
      if ((lower == mostRecent[1]) && (upper == mostRecent[3])) {
        let lowerwins = add ? Number(mostRecent[2]) + Number(bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).getValue()) : Number(bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).getValue()) - Number(mostRecent[2]);
        let upperwins = add ? Number(mostRecent[4]) + Number(bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).getValue()) : Number(bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).getValue()) - Number(mostRecent[4]);
        bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).setValue(lowerwins);
        bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).setValue(upperwins);
        if ((lowerwins >= gamesToWin) && (lowerwins > upperwins)) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[1]);
          bracket.getRange(i*(2**(j+1)) - offset, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if ((upperwins >= gamesToWin) && (upperwins > lowerwins)) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[3]);
          bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if (!add) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue('');
          bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j, 1, 2).setBackground(null);
        }
        found = true;
        break;
      } else if ((lower == mostRecent[3]) && (upper == mostRecent[1])) {
        let lowerwins = add ? Number(mostRecent[4]) + Number(bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).getValue()) : Number(bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).getValue()) - Number(mostRecent[4]);
        let upperwins = add ? Number(mostRecent[2]) + Number(bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).getValue()) : Number(bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).getValue()) -  Number(mostRecent[2]);
        bracket.getRange(i*(2**(j+1)) - offset, 2*j+1).setValue(lowerwins);
        bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j+1).setValue(upperwins);
        if ((lowerwins >= gamesToWin) && (lowerwins > upperwins)) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[3]);
          bracket.getRange(i*(2**(j+1)) - offset, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if ((upperwins >= gamesToWin) && (upperwins > lowerwins)) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[1]);
          bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if (!add) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue('');
          bracket.getRange(i*(2**(j+1)) - offset - 2**j, 2*j, 1, 2).setBackground(null);
        }
        found = true;
        break;
      }
    }
    if (found) {
      if (add) {
        sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name);
        sendResultToDiscord('Bracket', bracket.getSheetId());
      } else {
        sheet.getSheetByName('Results').deleteRow(results.length);
      }
      break;
    }
  }
}

//Swiss system functions

function swissOptions(stage) {
  return [
    ["''" + stages[stage].name + "' Swiss Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""],
    ["Seeding Method", "standard"]
  ];
}

function createSwiss(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var swiss = sheet.insertSheet(stages[stage].name + ' Standings');
  swiss.deleteRows(101, 900);
  swiss.setColumnWidth(1, 20);
  swiss.setColumnWidth(2, 150)
  swiss.setColumnWidths(3, 2, 50);
  swiss.setColumnWidth(6, 150);
  swiss.setColumnWidths(7, 2, 20);
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
  processing.deleteRows(101, 900);
  processing.getRange('E1').setValue('-');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var players = sheet.getSheetByName(options[1][1]).getRange('B2:B' + (Number(options[2][1]) + 1)).getValues().flat();
  
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
  switch (options[5][1].toLowerCase()) {
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
    processing.appendRow([playerToInsert, 'BYE', options[4][1], 0]);
    processing.appendRow(['BYE', playerToInsert, 0, options[4][1]]);
  }
}

function updateSwiss(stage, add) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  var swiss = sheet.getSheetByName(stages[stage].name + ' Standings');
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  
  //if removing we want to find the last processing row
  if (!add) {
    let flatResults = processing.getRange('A:A').getValues();
    var frlength = flatResults.length;
    for (let i=1; i<frlength; i++) {
      if (!flatResults[i][0]) {
        frlength = i;
      }
    }
  }
  
  //update matches and standings - standings updates only occur if the reported match was in the match list
  var roundMatches = swiss.getRange(2, 6, Math.ceil(options[2][1]/2),4).getValues();
  var found = false;
  for (let i=0; i < roundMatches.length; i++) {
    if ((roundMatches[i][0] == mostRecent[1]) && (roundMatches[i][3] == mostRecent[3])) {
      var leftwins = add ? Number(roundMatches[i][1]) + Number(mostRecent[2]) : Number(roundMatches[i][1]) - Number(mostRecent[2]);
      var rightwins = add ? Number(roundMatches[i][2]) + Number(mostRecent[4]) : Number(roundMatches[i][2]) - Number(mostRecent[4]);
      found = true;
    } else if ((roundMatches[i][0] == mostRecent[3]) && (roundMatches[i][3] == mostRecent[1])) {
      var leftwins = add ? Number(roundMatches[i][1]) + Number(mostRecent[4]) : Number(roundMatches[i][1]) - Number(mostRecent[4]);
      var rightwins = add ? Number(roundMatches[i][2]) + Number(mostRecent[2]) : Number(roundMatches[i][2]) - Number(mostRecent[2]);
      found = true;
    }    
    if (found) {
      swiss.getRange(i+2, 7, 1, 2).setValues([[leftwins, rightwins]]);
      if (add) { 
        processing.appendRow([mostRecent[1], mostRecent[3], mostRecent[2], mostRecent[4]]);
        processing.appendRow([mostRecent[3], mostRecent[1], mostRecent[4], mostRecent[2]]);
        sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name);
        sendResultToDiscord('Standings', swiss.getSheetId());
        //make next round's match list if needed
        newSwissRound(stage);
      } else {
        sheet.getSheetByName('Results').deleteRow(results.length);
        processing.getRange(frlength-1, 1, 2, 4).setValues([['','','',''],['','','','']]);
      }
      break;
    }
  }
  
  //remove current round if we are removing result and it wasn't there, then try again on previous round
  if (!found && !add) {
    swiss.deleteColumns(5, 5);
    swiss.setColumnWidth(5, 100);
    if (options[2][1] % 2) {
      processing.getRange(frlength-1, 1, 2, 4).setValues([['','','',''],['','','','']]);
    }
    updateSwiss(stage, false);
  }
}

function newSwissRound(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var swiss = sheet.getSheetByName(stages[stage].name + ' Standings');
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var roundInfo = swiss.getRange('G1:H1').getValues()[0];
  
  if ((roundInfo[1] == 1) && (Number(roundInfo[0]) < options[3][1])) {
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
    swiss.setColumnWidths(7, 2, 20);
    swiss.setColumnWidth(9, 150);
    swiss.getRange('F1').setHorizontalAlignment('right');
    swiss.getRange('G1').setHorizontalAlignment('left');
    swiss.getRange('H1').setFontColor('#ffffff');
    swiss.getRange('I:I').setHorizontalAlignment('right');
    swiss.getRange('J:J').setBackground('#707070');
    swiss.getRange(1, 6, roundMatches.length, 4).setValues(roundMatches);
    
    if (byeplace) {
      let byeIdx = assigned.indexOf('BYE');
      if (byeIdx % 2) {
        processing.appendRow([assigned[byeIdx-1],'BYE', options[4][1], 0]);
        processing.appendRow(['BYE', assigned[byeIdx-1], 0, options[4][1]]);
      } else {
        processing.appendRow([assigned[byeIdx+1],'BYE', options[4][1], 0]);
        processing.appendRow(['BYE', assigned[byeIdx+1], 0, options[4][1]]);
      }
    }
  }
}

//Preset Random functions

function randomOptions(stage) {
  return [
    ["''" + stages[stage].name + "' Preset Random Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Number of Rounds", ""],
    ["Games per Match", ""]
  ];
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
  var players = sheet.getSheetByName(options[1][1]).getRange('B2:B' + (Number(options[2][1]) + 1)).getValues().flat();
  if (options[2][1] % 2) {players.push('OPEN');}
  var pllength = players.length;
  players.sort(lowerCaseOrder);
  
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
  processing.deleteRows(101, 900);
  processing.getRange('D1:D2').setValues([['-'],['-']]);
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
    processing.appendRow([byePlayer, options[4][1], 0]);
    processing.appendRow(['BYE', 0, options[4][1]]);
  }
  
  //standings table copy from processing
  var table = [['Pl', 'Player', 'Win Pct','Wins', 'Losses', 'SoS']];
  for (let i=1; i <= options[2][1]; i++) {
    table.push([i, "='"+stages[stage].name+" Processing'!E"+(i+1), "='"+stages[stage].name+" Processing'!F"+(i+1),
      "='"+stages[stage].name+" Processing'!G"+(i+1), "='"+stages[stage].name+" Processing'!H"+(i+1), "='"+stages[stage].name+" Processing'!I"+(i+1)]);
  }
  standings.getRange(1,1,table.length,6).setValues(table);
}

function updateRandom(stage, add) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var standings = sheet.getSheetByName(stages[stage].name + ' Standings');
  var matches = standings.getRange(2, 8, options[2][1], 2*options[3][1]+1).getValues();
  var mlength = matches.length;
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  var found = false;
  
  for (let i=0; i<mlength; i++) {
    if (matches[i][0] == mostRecent[1]) {
      for (let j=0; j < options[3][1]; j++) {
        if (matches[i][2*j+1] == mostRecent[3]) {
          var p1wins = add ? Number(matches[i][2*j+2]) + Number(mostRecent[2]) : Number(matches[i][2*j+2]) - Number(mostRecent[2]);
          var cell1 = [i+2,9+2*j];
          found = true;
        }
      }
    } else if (matches[i][0] == mostRecent[3]) {
      for (let j=0; j < options[3][1]; j++) {
        if (matches[i][2*j+1] == mostRecent[1]) {
          var p2wins = add ? Number(matches[i][2*j+2]) + Number(mostRecent[4]) : Number(matches[i][2*j+2]) - Number(mostRecent[4]);
          var cell2 = [i+2,9+2*j];
        }
      }
    }
  }
  if (found) {
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
    if (add) {
      processing.appendRow([mostRecent[1], mostRecent[2], mostRecent[4]]);
      processing.appendRow([mostRecent[3], mostRecent[4], mostRecent[2]]);
      sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name);
      sendResultToDiscord('Standings', standings.getSheetId());
    } else {
      //find the last processing row
      let flatResults = processing.getRange('A:A').getValues();
      let frlength = flatResults.length;
      for (let i=2+options[2][1]+((options[2][1]*options[3][1]) % 2); i<frlength; i++) {
        if (!flatResults[i][0]) {
          frlength = i;
        }
      }
      processing.getRange(frlength-1, 1, 2, 3).setValues([['','',''],['','','']]);
      sheet.getSheetByName('Results').deleteRow(results.length);
    }
  }
}

//Groups functions

function groupsOptions(stage) {
  return [
    ["''" + stages[stage].name + "' Preset Random Options", ""],
    ["Seeding Sheet", ""],
    ["Number of Players", ""],
    ["Players per Group", ""],
    ["Grouping Method", "pool draw"],
    ["Aggregation Method", "rank then wins"],
    ["Games per Match", ""]
  ];
}

function createGroups(stage) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var standings = sheet.insertSheet(stages[stage].name + ' Standings');
  standings.deleteColumns(6, 21);
  standings.setColumnWidth(1, 150);
  standings.setColumnWidths(2, 4, 50);
  var processing = sheet.insertSheet(stages[stage].name + ' Processing');
  var players = sheet.getSheetByName(options[1][1]).getRange('B2:B' + (Number(options[2][1]) + 1)).getValues().flat();
  var ngroups = Math.ceil(options[2][1]/options[3][1]);
  
  //establish groups
  var groups = [];
  var openCounter = 0;
  switch (options[4][1].toLowerCase()) {
    case 'random':
      shuffle(players);
      while (players.length % options[3][1]) {players.push('OPEN'+ ++openCounter);}
      break;
    case 'pool draw':
      while (players.length % options[3][1]) {players.push('OPEN'+ ++openCounter);}
      for (let i=0; i<options[3][1]; i++) {
        shuffle(players, ngroups*i, ngroups*(i+1)-1);
      }
    default:
      while (players.length % options[3][1]) {players.push('OPEN'+ ++openCounter);}
  }
  var pllength = players.length;
  for (let i=0; i<pllength; i++) {
    if (Math.floor(i/ngroups) % 2) {
      groups.push(/^OPEN\d+$/.test(players[i]) ? [players[i], ngroups-1-(i%ngroups), -1, 2] : [players[i], ngroups-1-(i%ngroups), 0, 0]);
    } else {
      groups.push(/^OPEN\d+$/.test(players[i]) ? [players[i], i%ngroups, -1, 2] : [players[i], i%ngroups, 0, 0]); 
    }
  }
  groups.sort((a,b) => lowerCaseOrder(a[0],b[0]));
  processing.getRange(1, 1, groups.length, 4).setValues(groups);
  processing.getRange('F1').setFormula('=query(A:D, "select A, avg(B), sum(C), sum(C)+sum(D) where not A=\'\' group by A")');
  processing.getRange('J2').setFormula('=arrayformula(iferror(H2:H' + (pllength + 1) + '/I2:I' + (pllength + 1) + ',0))');
  processing.getRange('L1').setFormula('=query(F:J, "select F, G, I/' + options[6][1] + ', H, I-H, J where not F=\'\' order by G, J desc")');
  var ranks = [];
  for (let i=0; i<ngroups; i++) {
    for (let j=1; j<=options[3][1]; j++) {
      ranks.push([j]);
    }
  }
  processing.getRange(2, 18, ranks.length).setValues(ranks);
  
  //make standings
  var tables = [];
  var groupCounter = 1;
  for (let i=0; i<pllength; i++) {
    if (i % options[3][1] == 0) {
      tables.push(['','','','','']);
      tables.push(['Group ' + groupCounter++, '', '', '', '']);
      standings.getRange(tables.length, 1, 1, 5).merge().setHorizontalAlignment('center').setBackground('#bdbdbd');
      tables.push(['Player','Played','Wins','Losses','Win Pct']);
    }
    if (i % options[3][1] == options[3][1]-1) {
      tables.push(["=IF(regexmatch('" + stages[stage].name + " Processing'!L" + (i+2) + ",\"^OPEN\\d+$\"),, '" + stages[stage].name + " Processing'!L" + (i+2)+ ")",
        "=IF(regexmatch('" + stages[stage].name + " Processing'!L" + (i+2) + ",\"^OPEN\\d+$\"),, '" + stages[stage].name + " Processing'!N" + (i+2)+ ")",
        "=IF(regexmatch('" + stages[stage].name + " Processing'!L" + (i+2) + ",\"^OPEN\\d+$\"),, '" + stages[stage].name + " Processing'!O" + (i+2)+ ")",
        "=IF(regexmatch('" + stages[stage].name + " Processing'!L" + (i+2) + ",\"^OPEN\\d+$\"),, '" + stages[stage].name + " Processing'!P" + (i+2)+ ")",
        "=IF(regexmatch('" + stages[stage].name + " Processing'!L" + (i+2) + ",\"^OPEN\\d+$\"),, '" + stages[stage].name + " Processing'!Q" + (i+2)+ ")"
      ]);
      tables.push(['','','','','']);
    } else {
      tables.push(["='" + stages[stage].name + " Processing'!L" + (i+2), "='" + stages[stage].name + " Processing'!N" + (i+2),
        "='" + stages[stage].name + " Processing'!O" + (i+2), "='" + stages[stage].name + " Processing'!P" + (i+2), "='" + stages[stage].name + " Processing'!Q" + (i+2)]);
    }
  }
  standings.getRange(1, 1, tables.length, 5).setValues(tables);
  
  //make aggregates
  switch (options[5][1].toLowerCase()) {
    case 'win only':
      processing.getRange('T1').setFormula('=query(L:R, "select L,R,M,Q where not L=\'\' order by Q desc")');
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Win Percentage']];
      for (let i=0; i<options[2][1]; i++) {
        aggTable.push([i+1, "='" + stages[stage].name + " Processing'!T" + (i+2), "='" + stages[stage].name + " Processing'!U" + (i+2),
          "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!V" + (i+2) + ")",
          "='" + stages[stage].name + " Processing'!W" + (i+2)
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 5).setValues(aggTable);
      break;
    case 'separate':
      processing.insertColumnsAfter(20, 4*options[3][1]);
      for (let i=0; i<options[3][1]; i++) {
        processing.getRange(1, 20+4*i).setFormula('=query(L:R, "select L,R,M,Q where R='+(i+1)+' and not L=\'\' order by Q desc")');
        let aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation ' + (i+1) + ((i==0) ? 'sts' : (i==1) ? 'nds' : (i==2) ? 'rds' : 'ths'));
        let aggTable = [['Seed', 'Player', 'Place', 'Group', 'Win Percentage']];
        for (let j=0; j<ngroups; j++) {
          if (i == options[3][1]-1) {
            let openCheckString = "=IF(regexmatch('" + stages[stage].name + " Processing'!" + columnFinder(20+4*i) + (j+2) + ",\"^OPEN\\d+$\"),,";
            aggTable.push(["=IF(regexmatch('" + stages[stage].name + " Processing'!" + columnFinder(20+4*i) + (j+2) + ",\"^OPEN\\d+$\"),," + (j+1) + ")",
              openCheckString + "'" + stages[stage].name + " Processing'!" + columnFinder(20+4*i) + (j+2) + ")",
              openCheckString + "'" + stages[stage].name + " Processing'!" + columnFinder(21+4*i) + (j+2) + ")",
              openCheckString + "index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!" + columnFinder(22+4*i) + (j+2) + "))",
              openCheckString + "'" + stages[stage].name + " Processing'!" + columnFinder(23+4*i) + (j+2) + ")"
            ]);
          } else {
            aggTable.push([j+1,
              "='" + stages[stage].name + " Processing'!" + columnFinder(20+4*i) + (j+2),
              "='" + stages[stage].name + " Processing'!" + columnFinder(21+4*i) + (j+2),
              "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!" + columnFinder(22+4*i) + (j+2) + ")",
              "='" + stages[stage].name + " Processing'!" + columnFinder(23+4*i) + (j+2)
            ]);
          } 
        }
        aggregate.getRange(1, 1, aggTable.length, 5).setValues(aggTable);
      }
      break;
    case 'fixed':
      processing.insertColumnsAfter(20, 4);
      processing.getRange(1, 20).setFormula('=query(L:R, "select L,R,M,Q where R=1 and not L=\'\' order by Q desc")');
      processing.getRange(1, 24).setFormula('=query(L:R, "select L,R,M,Q where R=2 and not L=\'\' order by Q desc")');
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Win Percentage']];
      for (let i=0; i<ngroups; i++) {
        aggTable.push([i+1, "='" + stages[stage].name + " Processing'!T" + (i+2), "='" + stages[stage].name + " Processing'!U" + (i+2),
          "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!V" + (i+2) + ")",
          "='" + stages[stage].name + " Processing'!W" + (i+2)
        ]);
      }
      for(let i=ngroups; i>0; i-=2) {
        aggTable.push([2*ngroups+1-i, "='" + stages[stage].name + " Processing'!X" + i, "='" + stages[stage].name + " Processing'!Y" + i,
          "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!Z" + i + ")",
          "='" + stages[stage].name + " Processing'!AA" + i
        ]);
        aggTable.push([2*ngroups+2-i, "='" + stages[stage].name + " Processing'!X" + (i+1), "='" + stages[stage].name + " Processing'!Y" + (i+1),
          "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!Z" + (i+1) + ")",
          "='" + stages[stage].name + " Processing'!AA" + (i+1)
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 5).setValues(aggTable);
      break;
    case 'none':
      break;
    default:
      processing.insertColumnsAfter(20, 4*options[3][1]);
      for (let i=0; i<options[3][1]; i++) {
        processing.getRange(1, 20+4*i).setFormula('=query(L:R, "select L,R,M,Q where R='+(i+1)+' and not L=\'\' order by Q desc")');
      }
      var aggregate = sheet.insertSheet(stages[stage].name + ' Aggregation');
      var aggTable = [['Seed', 'Player', 'Place', 'Group', 'Win Percentage']];
      for (let i=0; i<options[2][1]; i++) {
        let colOffset = 4*Math.floor(i/ngroups);
        let rowNum = (i%ngroups)+2
        aggTable.push([i+1,
          "='" + stages[stage].name + " Processing'!" + columnFinder(20+colOffset) + rowNum,
          "='" + stages[stage].name + " Processing'!" + columnFinder(21+colOffset) + rowNum,
          "=index('" + stages[stage].name + " Standings'!A:A, 2+" + (4 + Number(options[3][1])) + "*'" + stages[stage].name + " Processing'!" + columnFinder(22+colOffset) + rowNum + ")",
          "='" + stages[stage].name + " Processing'!" + columnFinder(23+colOffset) + rowNum
        ]);
      }
      aggregate.getRange(1, 1, aggTable.length, 5).setValues(aggTable);
  }
}

function updateGroups(stage, add) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  var options = sheet.getSheetByName('Options').getRange(stage*10 + 1,1,10,2).getValues();
  var processing = sheet.getSheetByName(stages[stage].name + ' Processing');
  var groups = processing.getRange(1, 1, options[3][1]*Math.ceil(options[2][1]/options[3][1]), 2).getValues();
  
  if (mostRecent[1] == mostRecent[3]) {return;}
  
  if (add) {
    for (let i=0; i<options[2][1]; i++) {
      if (groups[i][0] == mostRecent[1]) {
        var p1Group = groups[i][1];
      } else if (groups[i][0] == mostRecent[3]) {
        var p2Group = groups[i][1];
      }
    }
    if (p1Group == p2Group) {
      sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name);
      processing.appendRow([mostRecent[1], p1Group, mostRecent[2], mostRecent[4]]);
      processing.appendRow([mostRecent[3], p1Group, mostRecent[4], mostRecent[2]]);
      sendResultToDiscord('Standings',sheet.getSheetByName(stages[stage].name + ' Standings').getSheetId());
    }
  } else {
    sheet.getSheetByName('Results').deleteRow(results.length);
    let flatResults = processing.getRange('A:A').getValues();
    let frlength = flatResults.length;
    for (let i=groups.length+1; i<frlength; i++) {
      if (!flatResults[i][0]) {
        frlength = i;
      }
    }
    processing.getRange(frlength-1, 1, 2, 4).setValues([['','','',''],['','','','']]);
  }
}