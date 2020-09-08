// A name and type for each stage, ie [{name: 'Swiss', type: 'swiss}, {name: 'Elimination', type: 'single'}]
var stages = [{name: 'Swiss', type: 'swiss'}];

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
        createSwiss(i);
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
      updateSwiss(lastResultStage, false);
      break;
    default:
      sheet.getSheetByName('Results').deleteRow(results.length);
  }
}

function testWebhook() {
  var admin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];
  
  var possibles = ["[Hi! I'm your talking ticket](https://www.youtube.com/watch?v=9-QWW2Cxn98).",
                   "[Ahh, technology](https://www.youtube.com/watch?v=OZT9IvH5mC8)."];
  
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
  names.sort();
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
        var opts = swissOptions(i);
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
          updateSwiss(i, true);
      }
    }
  }
}

function sendResultToDiscord(displayType) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var admin = sheet.getSheetByName('Administration').getDataRange().getValues();
  var webhook = admin[3][1];
  var results = sheet.getSheetByName('Results').getDataRange().getValues();
  var mostRecent = results[results.length - 1];
  
  var display = '**' + mostRecent[1] + ' ' + mostRecent[2] + '-' + mostRecent[4] + ' ' + mostRecent[3] + '**';
  if (mostRecent[5]) {
    display += '\n' + mostRecent[5];
  }
  display += '\n**[Current ' + displayType + '](' + admin[1][1] + ')**';
  
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
      bracket.setColumnWidth(j, 40);
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
        if (lowerwins >= gamesToWin) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[1]);
          bracket.getRange(i*(2**(j+1)) - offset, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if (upperwins >= gamesToWin) {
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
        if (lowerwins >= gamesToWin) {
          bracket.getRange(i*(2**(j+1)) - offset - 2**(j-1), 2*j+2).setValue(mostRecent[3]);
          bracket.getRange(i*(2**(j+1)) - offset, 2*j, 1, 2).setBackground('#D9EBD3');
        } else if (upperwins >= gamesToWin) {
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
        sendResultToDiscord('Bracket');
      } else {
        sheet.getSheetByName('Results').deleteRow(results.length);
      }
      break;
    }
  }
}

//Swiss functions

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
  swiss.setColumnWidths(7, 2, 40);
  swiss.setColumnWidth(9, 150);
  swiss.setFrozenColumns(4);
  swiss.setFrozenRows(1);
  swiss.getRange('F1').setHorizontalAlignment('right');
  swiss.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=ISBLANK(B1)').setFontColor('#FFFFFF').setRanges([swiss.getRange('A:A')]).build()]);
  var processing = sheet.insertSheet(stages[stage].name + ' Processing');
  processing.deleteRows(101, 900);
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
    let flatResults = processing.getRange('A:D').getValues();
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
        sendResultToDiscord('Standings');
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
    swiss.setColumnWidths(7, 2, 40);
    swiss.setColumnWidth(9, 150);
    swiss.getRange('F1').setHorizontalAlignment('right');
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