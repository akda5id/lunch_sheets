# Lunch (in the) Sheets
Some Apps Script tooling to bring your data from [Lunch Money](https://lunchmoney.app/?refer=be4tew9v) to Google Sheets.
* Import all transactions to a sheet called "LM-Transactions", and update it with new transactions, semi-intelligently.
* Sum up category and tag totals per day, and per month.
* Track account totals and net worth over time.
* A user function to total a category or tag for a date range.

### Install:
1. Open the sheet you would like to use the tools in. Choose "Apps Script" from the Extensions menu. That will open a code editing page, if you don't have any existing Apps Scripts, it will be an empty function called "myFunction" in a file called Code.gs. Delete that empty function, and copy and paste everything [from Code.js](https://raw.githubusercontent.com/akda5id/lunch_sheets/main/Code.js) in this repo into your Code.gs.
1. If you do have other code in there, you can create a new file, and put this code into it. Heads up if you are using a custom menu, or otherwise using onLoad, there will be only be one onLoad called, so you will have to sort that out (put my onLoad code into yours, probably.)
1. Give the project a name (change "Untitled Project" to "Lunch Money Script", or whatever you want), then click the save icon, and close the window. 
1. Reload your spreadsheet and you should see a menu "Lunch Money" appear. Choose the "Set API Key" option. You will get a permission warning at this point. Click through that, it's saying that you are giving yourself permission to access your own data. On the "Google hasn’t verified this app" page, click "advanced", then "Go to…". 
1. After that completes, choose "Set API Key" again, to actually do it. Then put in your API key ([get one here](https://my.lunchmoney.app/developers)) at the prompt.

### Usage:
* Choose "Load Transactions" from the Lunch Money menu to load your transactions. On first run it will get all transactions. After the first run it will check for updated transactions in the month previous, and add new ones up until today to the end of the sheets.
* Custom function `LMCATTOTAL` sums a category (or tag) over months, function call looks like: `LMCATTOTAL(category, startDate, endDate, random)` for example: `=LMCATTOTAL("Travel", "2023-10", "2024-12", 'LM-Transactions'!R1)` 'LM-Transactions'!R1 is an incrementing counter, if LMWriteRandom is true in the settings (see below). Google sheets custom functions are quite slow, so recommend you don't use too much of this. You can run your own SUM's over ranges in the Months sheet, which will be much faster.

### Settings:
In the Lunch Money menu, there is an option to "Update Categories". Run this if you change the name, grouping, or hide from budget, income, etc. of categories. I cache them between runs as I don't expect them to change often, so it's a manual update. Note that if you do adjust them, it will only update for transactions in LookbackMonths (see below). If you want to rerun on all transactions, just delete or rename the LM-\* sheets.

There are a few things at the top of [Code.gs](Code.js) that you can adjust:

LMdebug: Write tracing info to the apps script log, if you are having trouble with the script, this might help us figure out what's going on.

LMJumpOnFinish: If you want it to take you to the end of the transaction sheet after an update is run. False will leave you wherever you were when you ran the update.

LMTransactionsLookbackMonths: Number of full months we will pull transactions from, prior to the current one, to check for updated category, etc. This one you should keep tight if you can, for efficiency. Basically if you always have all your transactions perfectly set and organized in Lunch Money before you pull them into sheets, you could set this to 0. But if you want to pull things into sheets during the month, and you only go through and categorize in Lunch Money at the end of the month, it should be at least 1, to make sure you pull in updates from the previous month. If you don't have so many transactions a month, it's fine to leave it at 2, then you don't really have to think about it.

LMTransactionsLookback: Max number of transactions you would ever get from today to the date resulting from LMTransactionsLookbackMonths. So if you pull once a month, and have a one month lookback, you need to have at least two months of transactions here. Be generous, it's fast. Script will throw an error if it's ever too little.

LMCoalesce\*: If you want to write the various sum up category and tags sheets. Should be self explanatory with the comments there.

LMTrack\*: If you want to track account totals and net worth.

LMWriteRandom: Provides an incrementing counter in `'LM-Transactions'!R1` so that Custom Functions run after you update (Google sheets caches function results). Also fun to see how many updates you have run :)

### Help:
There is a thread on the [Lunch Money Discord](https://discord.com/channels/842337014556262411/1176857773925998642), let me know how things are going. If you are not joined to the discord already [here is the signup link](https://discord.gg/vSz6jjZuj8).

### Random Notes:
* API key is per spreadsheet, so you can use separate "budgets" if you want, as long as you don't need them in the same spreadsheet.
* You can re-order columns in the LM-Days and LM-Months sheets as you like (except date, leave that at column A!). LM-Accounts you can reorder all except A and B. You can't re-order LM-Transactions, but you can hide columns.
* You should make sure the Spreadsheet timezone is set to what you expect. (File -> Settings)

### To Do:
* Expose a way to update Accounts (for name changes and such).
* Better error handling and edge case detection. It's working for me now, but I expect to find places where it breaks as I use it more, and hear about many more as other people start using it.
* Perhaps add functionality to `LMCATTOTAL` such as working by days also.

### Security and Privacy Notes:
This script is read only on Lunch Money, as you can see in apiRequest, method is hard coded to 'get'. Hopefully one day Jen creates read only API keys so you can enforce that on your scripts that way.

I only call the Lunch Money API, but the security warning you get when you first run the script warns about "Create[ing] a network connection to any external service". If you would like to make sure that I can't slip anything sneaky past you, you can add these lines to your [manifest](https://developers.google.com/apps-script/concepts/manifests):

	"oauthScopes": [
  		"https://www.googleapis.com/auth/script.external_request",
  		"https://www.googleapis.com/auth/spreadsheets.currentonly"
	],
	"urlFetchWhitelist": [
  		"https://dev.lunchmoney.app/"
	],

This is not necessary, only helps if I turn evil, and you manually update the script to my new evil version, since there is no auto update functionality.