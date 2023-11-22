# Lunch (in the) Sheets
Some Apps Script tooling to bring your data from [Lunch Money](https://lunchmoney.app/?refer=be4tew9v) to Google Sheets.
* (Working, beta) Import all transactions to a sheet called "LM-Transactions-All", and update it with new transactions, semi-intelligently.
* (Planned, not started) Coalesce category and tag totals per day, and per month.
* (Planned, not started) A user function to total a category or tag for a date range.
* (Planned, not started) Functions to return account totals and calculate net worth.

### Install:
* Open the sheet you would like to use the tools in. Choose "Apps Script" from the Extensions menu. That will open a code editing page, if you don't have any existing Apps Scripts, it will be an empty function called "myFunction" in a file called Code.gs. Delete that empty function, and copy and paste everything [from Code.js in this repo](https://raw.githubusercontent.com/akda5id/lunch_sheets/main/Code.js) into your Code.gs.
* If you do have other code in there, you can create a new file, and put this code into it. Heads up if you are using a custom menu, or otherwise using onLoad, there will be only be one onLoad called, so you will have to sort that out (put my onLoad code into yours, probably.)
* Give the project a name (change "Untitled Project" to "Lunch Money Script", or whatever you want), then click the save icon, and close the window. Reload your spreadsheet and you should see a menu "Lunch Money" appear. Choose the "Set API Key" option. You will get a permission warning at this point. Click through that, it's saying that you are giving yourself permission to access your own data. On the "Google hasn’t verified this app" page, click "advanced", then "Go to…". After that completes, choose "Set API Key" again, to actually do it. Then put in your API key ([get one here](https://my.lunchmoney.app/developers)) at the prompt.

### Usage:
* Choose "Load Transactions" from the Lunch Money menu to load your transactions. On first run it will get all transactions from the first transaction date you enter. After the first run it will look at transactions in the last 60 days, and update existing ones with data from Lunch Money, and add new ones to the end of the sheet.
* That's all it does for now, more coming!

### Help:
* Check to see if I have created a thread in the [Lunch Money Discord](https://discord.com/channels/842337014556262411/1134597088504729780), you can bug me there I guess.

##### Random Notes:
* API key is per spreadsheet, so you can use separate "budgets" if you want, as long as you don't need them in the same spreadsheet.

##### To Do:
* Expose a way to update Accounts (for name changes and such).
* Better error handling and edge case detection. It's working for me now, but I expect to find places where it breaks as I use it more, and hear about many more as other people start using it.

##### Security and Privacy Notes:
I only call the Lunch Money API, but the security warning you get when you first run the script warns about "Create[ing] a network connection to any external service". If you would like to make sure that I can't slip anything sneaky past you, you can add these lines to your [manifest](https://developers.google.com/apps-script/concepts/manifests):
`"oauthScopes": [
  	"https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets.currentonly"
  ],
  "urlFetchWhitelist": [
    "https://dev.lunchmoney.app/"
  ],`
  This is not necessary, only helps if I turn evil, and you manually update the script to my new evil version, since there is no auto update functionality.