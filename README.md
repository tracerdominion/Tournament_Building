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

The first 3 rows should be the following:
```
["''" + stages[stage].name + "' <stage type name> Options", ""],["Seeding Sheet", ""],["Number of Players", ""]
```

You may include preset values for some keys as suggestions. They do not need to match the default when you eventually use the option.

While not unacceptable, I recommend not asking for values used purely for aesthetics. Rather, inform the user in documentation that they may change these things as they wish.

Once completed, it should be appropriately added to the switch statement in the `addStages()` function.

#### Create functions

Should take a single parameter `stage` which is the index of the stage in the stages array at the very top of the script.
To match existing names, called `create<Stage_type>`.

This function should create and set up all sheets for this stage, without anything non-aesthetic needed to be added by the user. Be sure to consider how results will be processed.

Once completed, it should be appropriately added to the switch statement in the `setupSheets()` function.

#### Update (with new results) functions

Should take a parameter `stage` which is the index of the stage in the stages array at the very top of the script.
May take an additional parameter which tells it whether the update is an addition or removal. Since removals often require the same searches as additions, it can be useful to handle both in the same function.
To match existing names, called `update<Stage_type>`.

This should take the most recent result and place it where it needs to go to update the stage to include it. It should be able to handle partial matches.

You can access the most recent result through:
```
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var results = sheet.getSheetByName('Results').getDataRange().getValues();
var mostRecent = results[results.length - 1];
```

`mostRecent` will be an array of length 6. Relevant are the items at indexes 1-4:
* `mostRecent[1]` is the first player
* `mostRecent[2]` is wins for the first player (as a string - be careful)
* `mostRecent[3]` is the second player
* `mostRecent[4]` is wins for the second player (again as a string)

In addition to doing whatever is needed to add these results, if successfully added it should also:
* Append the stage name to the most recent result. This can be done with `sheet.getSheetByName('Results').getRange(results.length, 7).setValue(stages[stage].name)` assuming `sheet` is defined as above.
* Make a call to the function `sendResultToDiscord(displayType, gid)` to show the results embed. `displayType` should be a string that describes the output of the stage, for example `'Standings'` or `'Bracket'`. `gid` should be the id of that output sheet, accessed through `getSheetId()`([ref](https://developers.google.com/apps-script/reference/spreadsheet/sheet#getSheetId())).

Once completed, it should be appropriately added to the switch statement in the `onFormSubmit()` function.

#### Removal functions

Should take a parameter `stage` which is the index of the stage in the stages array at the very top of the script.
To this point in time all removal functions have been incorporated into the stage's Update function.

This should again take the most recent result, but in this case remove it. It can be assumed that the result has been already been added whenever this is called. It should also delete the last line of the Results sheet.

Once completed, it should be appropriately added to the switch statement in the `removeLastResult()` function.

#### Helpers

Right now there are two helper functions in the script that may be useful in stage creation:

`shuffle(array, start, end)` - randomly permutes the elements indexed between start and end, inclusive, of the array in place. This can be useful for seeding randomization.

`columnFinder(n)` - returns the letter code for the nth spreadsheet column (1 => 'A', 2 => 'B', ... 26 => 'Z', 27 => 'AA', ...) up to 702. This can be useful when adding spreadsheet formulas to your sheets. Please don't exceed 702 columns, whether you are using this function or not.