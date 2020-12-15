# Tournament_Building

Google Apps Script file for tournament building in Google Sheets.

This README is concerned with protocol for adding stage types. If you are looking to use this script to build a tournament, see the instructions at https://domi.link/tournament-building

### Adding a Stage Type

I am certainly willing to accept pull requests for new stage types and/or additional features on existing ones. Please familiarize yourself with existing ones before doing so to get an idea of what I am looking for from these, as well as [Google Apps Script Sheets Documentation](https://developers.google.com/apps-script/reference/spreadsheet/).

A couple of notes on what a stage type should look like:
* Should support any number of players - if a certain number is strongly desired, look at figuring out how byes might be incorporated
* Have one output sheet that provides all information about both player position and matches to be scheduled
* Take up no more than 3 sheets (except by option) - one for player position and matches, possibly one for processing, and possibly one (or optionally more) for a tranformation of the standings
* Should include documentation as seen in section 6 of the instructions. You may message this to me via Discord.

A stage type must include functions which can do the following:
* Return a key-value set of options
* Create and populate sheets for the stage to run
* Update those sheets when new results come in
* Remove the last result submitted

#### Options function

Should take a single parameter `stage` which is the index of the stage in the stages array at the very top of the script.
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

Should take a single parameter `stage` which is the index of the stage in the stages array at the very top of the script.
To match existing names, called `create<Stage_type>`.

This function should create and set up all sheets for this stage, without anything non-aesthetic needed to be added by the user. Be sure to consider how results will be processed.

Be sure to import players by index (not value or cell reference) from the seeding sheet - this is essential for automatic linking once results fill in.

Once completed, it should be appropriately added to the switch statements in both the `setupSheets()` and the `remakeStages()` functions.

#### Update (with new results) functions

Should take a parameter `stage` which is the index of the stage in the stages array at the very top of the script.
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

Should take parameters `stage`, which is the index of the stage in the stages array at the very top of the script, and `row`, which is the row in the Results sheet that is to be removed.
To match existing names, called `remove<Stage_type>`.

This should take the result in the `row` numbered row of the results sheet and attempt to remove it. If there are dependent matches with results (for example the match a winner plays in) it should fail to remove the result and log a failure message and reason. Otherwise, it should remove any sign of it from sheets intended for display. 
It should also remove the stage name label from that row in the results sheet, to indicate that the result is not being accounted for anywhere.

#### Helpers

Right now there are two helper functions in the script that may be useful in stage creation:

`shuffle(array, start, end)` - randomly permutes the elements indexed between start and end, inclusive, of the array in place. This can be useful for seeding randomization.

`columnFinder(n)` - returns the letter code for the nth spreadsheet column (1 => 'A', 2 => 'B', ... 26 => 'Z', 27 => 'AA', ...) up to 702. This can be useful when adding spreadsheet formulas to your sheets. Please don't exceed 702 columns, whether you are using this function or not.
