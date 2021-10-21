# Tournament_Building

Google Apps Script file for tournament building in Google Sheets.

This README is concerned with protocol for adding stage types. If you are looking to use this script to build a tournament, see the instructions at https://domi.link/tournament-building

### Adding a Stage Type

I am certainly willing to accept pull requests for new stage types and/or additional features on existing ones. Please familiarize yourself with existing ones before doing so to get an idea of what I am looking for from these, as well as [Google Apps Script Sheets Documentation](https://developers.google.com/apps-script/reference/spreadsheet/).

A couple of notes on what a stage type should look like:
* Should support any number of players - if a certain number is strongly desired, look at figuring out how byes might be incorporated
* Have one output sheet that provides all information about both player position and matches to be scheduled
* Take up no more than 3 sheets (except by option) - one for player position and matches, possibly one for processing, and possibly one (or optionally more) for a tranformation of the standings
* Should include documentation as seen in section A of the instructions. You may message this to me via Discord.

A stage type must include functions which can do the following:
* Return a key-value set of options
* Create and populate sheets for the stage to run
* Update those sheets when new results come in
* Remove a result with row given
* Break a tie between players on the player position sheet
* Drop a player from the stage (and adjust future matches accordingly)

#### Options function

Should take two parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage.
Should return an array of length 2 arrays (subsequently called 'rows') which are key-value pairs for user provided options. There should be no more than 9 total rows.
To match existing names, called `<stage_type>Options`.

This function should place key-value pairs for user provided options at the top of the options sheet. There should be no more than 8 options, for no more than 9 total rows used.

The first 3 rows should be the following:
```
["''" + stages[stage].name + "' <stage type name> Options", ""],["Seeding Sheet", ""],["Number of Players", ""]
```

You may include preset values for some keys as suggestions. Values which should be one of a number of options should have data validation with dropdown placed on the corresponding cells (see the Single Elimination Seeding Method option as an example).

While not unacceptable, I recommend not asking for values used purely for aesthetics. Rather, inform the user in documentation that they may change these things as they wish.

Once completed, it should be appropriately added to the switch statement in the `addStages()` function.

#### Create functions

Should take two parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage.
To match existing names, called `create<Stage_type>`.

This function should create and set up all sheets for this stage and add warnings upon edit where appropriate, without anything non-aesthetic needed to be added by the user. Be sure to consider how results will be processed.

Users should be alerted if an option they provide does not fit the stage.

Be sure to import players by index (not value or cell reference) from the seeding sheet - this is essential for automatic linking once results fill in.

Once completed, it should be appropriately added to the switch statement in the `addStage()` function. There should also be a process for deleting all associated sheets, which should be inserted into the switch statement in the `removeStage()` function.

#### Update (with new results) functions

Should take two parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage.
To match existing names, called `update<Stage_type>`.
Should return true if an update took place, and flase otherwise.

This should take the most recent result (from the form, not the spreadsheet) and place it where it needs to go to update the stage to include it. It should be able to handle partial matches.

You can access the most recent result through:
```
var form = FormApp.openByUrl(sheet.getSheetByName('Administration').getDataRange().getValues()[2][1]).getResponses();
var mostRecent = form[form.length - 1].getItemResponses().map(function(resp) {return resp.getResponse();});
```

`mostRecent` will be an array of length 6. Relevant are the items at indexes 1-4:
* `mostRecent[0]` is the first player
* `mostRecent[1]` is wins for the first player (as a string - be careful)
* `mostRecent[2]` is the second player
* `mostRecent[3]` is wins for the second player (again as a string)

In addition to doing whatever is needed to add these results, if successfully added it should also:
* Append the stage name to the most recent result. This can be done with `sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name)` assuming `sheet` is defined as above.
* Append location data to the most recent result for easier removal. This can be done with `results.getRange(results.getDataRange().getNumRows(), 8).setFontColor('#ffffff').setValue(<location_data>)` assuming sheet is defined as above.
* Make a call to the function `sendResultToDiscord(displayType, gid)` to show the results embed. `displayType` should be a string that describes the output of the stage, for example `'Standings'` or `'Bracket'`. `gid` should be the id of that output sheet, accessed through `getSheetId()`([ref](https://developers.google.com/apps-script/reference/spreadsheet/sheet#getSheetId())).
* Return true

Once completed, it should be appropriately added to the switch statement in the `onFormSubmit()` function.

#### Removal functions

Should take parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage, and `row`, which is the row in the Results sheet that is to be removed.
Should return true if the result is able to be removed, and false otherwise.
To match existing names, called `remove<Stage_type>`.

This should take the result in the `row` numbered row of the results sheet and attempt to remove it. If there are dependent matches with results (for example the match a winner plays in) it should fail to remove the result and alert a failure message and reason. Otherwise, it should remove any sign of it from sheets intended for display. 
It should also remove the stage name label from that row in the results sheet, to indicate that the result is not being accounted for anywhere.

Once completed, it should be appropriately added to the switch statement in the `onFormSubmit()` function.

#### Tiebreaking functions

Should take parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage, and `players`, which is an array of players involved in the tie ordered by what their standing should be (starting from the top of the tie)
Should return true if the tie is found and successfully broken.
To match existing names, called `tiebreak<Stage_type>`

This should rearrange the names of the players provided on a sheet giving ranks to the players to match the order provided. If the players are not in a tie, it should fail and alert the user. If the stage is not yet complete, it should ask for confirmation from the user.

Once completed, it should be appropriately added to the switch statement in the `addTiebreaker()` function.

#### Drop Player functions

Should take parameters `num` and `name`, the first of which is 8 less than the row number of the administration sheet corresponding to that stage, and the second of which is the name of the stage, and `player`, which is the name of the player to be dropped.
Does not need to return a value.
To match existing names, called `drop<Stage_type>`

This should ensure no future matches are assigned to the player, and drop them to the bottom of their standings. Other appropriate measures based on the nature of the stage type should be implemented.

Once completed, it should be appropriately added to the switch statement in the `dropPlayer()` function.

#### Helpers

Right now there are two helper functions in the script that may be useful in stage creation:

`shuffle(array, start, end)` - randomly permutes the elements indexed between start and end, inclusive, of the array in place. This can be useful for seeding randomization.

`columnFinder(n)` - returns the letter code for the nth spreadsheet column (1 => 'A', 2 => 'B', ... 26 => 'Z', 27 => 'AA', ...) up to 702. This can be useful when adding spreadsheet formulas to your sheets. Please don't exceed 702 columns, whether you are using this function or not.
