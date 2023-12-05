/**
 * @OnlyCurrentDoc
 */
"use strict";

function onOpen() {
 var menu = [{name: "Load Transactions", functionName: "updateTransactionsAll"},null,{name: "Update Categories", functionName: "updateCatagories"},{name: "Set API Key", functionName: "setApiKey"}];
 SpreadsheetApp.getActiveSpreadsheet().addMenu("Lunch Money", menu);
}

/**
 *  SETTINGS:
 */
const LMdebug = true;           //write tracing info to the apps script log

const LMJumpOnFinish = false;    //should we jump to the last row when we finish updating transactions?

const LMTransactionsLookbackMonths = 1; //number of full months we will pull transactions from, prior to the current one, to check
                                        //for updated category, etc. This one you should keep tight if you can, to reduce load on Lunch Money.

const LMTransactionsLookback = 1000;  //Max number of transactions you would ever get from today to LMTransactionsLookbackMonths.
                                      //Be generous, it's fast. Script will error with a warning if it is too small.

const LMCoalesce = true;       //Should we total up all categories and tags and write to separate sheet?
const LMCoalesceMonths = true; //if so, by months?
const LMCoalesceDays = true;   //and by days?

const LMTrackPlaidAccounts = true;   //Track plaid account values?
const LMTrackAssets = false;         //Track manually updated assets?

const LMWriteRandom = true;   //write an incrementing counter to 'LM-Transactions'!R1 to use in custom function calls to avoid caching.
/**
 *  END OF SETTINGS
 */

const LMDocumentProperties = PropertiesService.getDocumentProperties();
const LMActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const LMSpreadsheetTimezone = LMActiveSpreadsheet.getSpreadsheetTimeZone();
const LMScriptTimezone = Session.getScriptTimeZone();

function updateTransactionsAll() {
  var transactionsAllSheet = LMActiveSpreadsheet.getSheetByName("LM-Transactions");
  if (transactionsAllSheet == null) {
    let firstTransactionDate = '1970-01-01';
    var today = new Date();
    today = Utilities.formatDate(today, LMSpreadsheetTimezone, "yyyy-MM-dd");
    let transactions = loadTransactions(firstTransactionDate, today);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let {LMCategories, plaidAccountNames, assetAccountNames, plaidAccounts, assetAccounts} = loadCategoriesAndAccounts();
    let {parsedTransactions_2d, months, days} = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
    transactionsAllSheet = createTransactionsAllSheet();
    transactionsAllSheet.getRange(2, 1, parsedTransactions_2d.length, parsedTransactions_2d[0].length).setValues(parsedTransactions_2d);
    var transactionsAllLastRow = transactionsAllSheet.getLastRow();
    if (LMCoalesce) {writeCoalesed(months, days);}
    if (LMTrackPlaidAccounts || LMTrackAssets) {trackNW(plaidAccounts, plaidAccountNames, assetAccounts, assetAccountNames);}
  } else {
    var transactionsAllLastRow = transactionsAllSheet.getLastRow();
    let {LMCategories, plaidAccountNames, assetAccountNames, plaidAccounts, assetAccounts} = loadCategoriesAndAccounts();
    let {startDate, endDate} = calculateRelativeDates();
    let transactions = loadTransactions(startDate, endDate);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let {parsedTransactions_2d, months, days} = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
    let row = findIdTransactionsAll(parsedTransactions_2d[0][0].toFixed(0), transactionsAllSheet, transactionsAllLastRow);
    let transactionsLength = parsedTransactions_2d.length;
    if (transactionsLength >= LMTransactionsLookback) {throw new Error('CAUTION! LMTransactionsLookback is not large enough!');}
    if (transactionsAllSheet.getMaxRows() < row+transactionsLength+60) {
      let need = row+transactionsLength+500 - transactionsAllSheet.getMaxRows();
      if (LMdebug) {Logger.log('updateTransactionsAll: adding %s rows', need);}
      transactionsAllSheet.insertRowsAfter(transactionsAllLastRow, need);
    }
    if (LMdebug) {Logger.log('updateTransactionsAll: overwriting from row: %s', row);}
    transactionsAllSheet.getRange(row, 1, transactionsLength, parsedTransactions_2d[0].length).setValues(parsedTransactions_2d);
    transactionsAllSheet.getRange(row+transactionsLength, 1, 50, parsedTransactions_2d[0].length).clear();
    if (LMWriteRandom) {
      let range = transactionsAllSheet.getRange(1, 18, 1, 1);
      var foo = range.getValue();
      if (foo == '') { foo = 0; }
      foo += 1;
      range.setValue(foo);
    }
    if (LMCoalesce) {writeCoalesed(months, days);}
    if (LMTrackPlaidAccounts || LMTrackAssets) {trackNW(plaidAccounts, plaidAccountNames, assetAccounts, assetAccountNames);}
  }
  if (LMJumpOnFinish) {transactionsAllSheet.setActiveCell(transactionsAllSheet.getDataRange().offset(transactionsAllLastRow, 0, 1, 1));}
}

function findIdTransactionsAll(id, transactionsAllSheet, transactionsAllLastRow){
  var transactionsLookback = LMTransactionsLookback;
  let transactionAllIdsStart = transactionsAllLastRow - transactionsLookback;
  if (transactionAllIdsStart < 2) {
    transactionAllIdsStart = 2;
    transactionsLookback = transactionsAllLastRow + 1;
  }
  let transactionAllIds = transactionsAllSheet.getRange(transactionAllIdsStart, 1, transactionsLookback).getValues();

  let row = transactionAllIds.findIndex(foo => {return foo[0] == id});
  if (row == -1) { throw new Error('Didn\'t find transaction id: ' + id); } //TODO: Probably need to fall back to date?
  return row + transactionAllIdsStart;
}

function createTransactionsAllSheet() {
  var transactionsAllSheet = LMActiveSpreadsheet.insertSheet('LM-Transactions');
  let data = [['id', 'date', 'category name', 'payee', 'amount', 'notes', 'account name', 'tag', 'status', 'exclude from totals', 'exclude from budget', 'is income']]
  transactionsAllSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return transactionsAllSheet
}

function trackNW(plaidAccounts, plaidAccountNames, assetAccounts, assetAccountNames) {
  var netWorth = 0;
  var accountsSheet = LMActiveSpreadsheet.getSheetByName('LM-Accounts');
  if (accountsSheet == null) { accountsSheet = createSheet('LM-Accounts'); }
  var headers = accountsSheet.getRange(2, 1, 1, accountsSheet.getLastColumn()).getValues()[0];

  if (LMTrackPlaidAccounts) {
    for (const account of plaidAccounts) {
      let accountName = plaidAccountNames[account.id];
      var date = new Date(account.balance_last_update);
      date = Utilities.formatDate(date, LMSpreadsheetTimezone, "yyyy-MM-dd");
      var row = findDate(accountsSheet, date);
      if ( row == -1 ) { 
        row = accountsSheet.getLastRow() + 1;
        accountsSheet.getRange(row, 1).setValue(date);
      }
      if (accountsSheet.getRange(row,1).getValue() != date) {
        accountsSheet.insertRowBefore(row);
        accountsSheet.getRange(row, 1).setValue(date);
      }
      var amount = +(account.balance);
      if (account.type == 'credit') {
        if (amount > 0) {
          amount = 0 - amount;
        } else {
          amount = Math.abs(amount);
        }
      }
      netWorth += amount;
      let index = headers.indexOf(accountName);
      if ( index == -1 ) {
        index = headers.push(accountName) - 1; //push returns length, we want index which is one less
        accountsSheet.getRange(2, accountsSheet.getLastColumn()+1).setValue(accountName);
        accountsSheet.getRange(row,index+1).setValue(amount.toFixed(2));
        accountsSheet.getRange(1,index+1).setValue(amount.toFixed(2));
      } else {
        accountsSheet.getRange(row,index+1).setValue(amount.toFixed(2));
        accountsSheet.getRange(1,index+1).setValue(amount.toFixed(2));
      }
    }
  }

  if (LMTrackAssets) {
    for (const account of assetAccounts) {
      let accountName = assetAccountNames[account.id];
      if (account.balance_as_of == null) {
        continue;
      }
      var date = new Date(account.balance_as_of);
      date = Utilities.formatDate(date, LMSpreadsheetTimezone, "yyyy-MM-dd");
      var row = findDate(accountsSheet, date);
      if ( row == -1 ) { 
        row = accountsSheet.getLastRow() + 1;
        accountsSheet.getRange(row, 1).setValue(date);
      }
      if (accountsSheet.getRange(row,1).getValue() != date) {
        accountsSheet.insertRowBefore(row);
        accountsSheet.getRange(row, 1).setValue(date);
      }
      var amount = +(account.balance);
      if (account.type_name == 'credit' || account.type_name == 'loan' || account.type_name == 'other liability') {
        if (amount > 0) {
          amount = 0 - amount;
        } else {
          amount = Math.abs(amount);
        }
      }
      netWorth += amount;
      let index = headers.indexOf(accountName);
      if ( index == -1 ) {
        index = headers.push(accountName) - 1; //push returns length, we want index which is one less
        accountsSheet.getRange(2, accountsSheet.getLastColumn()+1).setValue(accountName);
        accountsSheet.getRange(row,index+1).setValue(amount.toFixed(2));
        accountsSheet.getRange(1,index+1).setValue(amount.toFixed(2));
      } else {
        accountsSheet.getRange(row,index+1).setValue(amount.toFixed(2));
        accountsSheet.getRange(1,index+1).setValue(amount.toFixed(2));
      }
    }
  }

  var today = new Date();
  today = Utilities.formatDate(today, LMSpreadsheetTimezone, "yyyy-MM-dd");
  var row = findDate(accountsSheet, today);
  if ( row == -1 ) { 
    row = accountsSheet.getLastRow() + 1;
    accountsSheet.getRange(row, 1).setValue(today);
  }
  if (accountsSheet.getRange(row,1).getValue() != today) {
    accountsSheet.insertRowBefore(row);
    accountsSheet.getRange(row, 1).setValue(today);
  }
  accountsSheet.getRange(row,2).setValue(netWorth.toFixed(2));
  accountsSheet.getRange(1,2).setValue(netWorth.toFixed(2));
}

function writeCoalesed(months, days) {
  if (LMCoalesceMonths) {
    var monthsSheet = LMActiveSpreadsheet.getSheetByName('LM-Months');
    if (monthsSheet == null) { monthsSheet = createSheet('LM-Months'); }
    var headers = monthsSheet.getRange(1, 1, 1, monthsSheet.getLastColumn()).getValues()[0];
    var monthKeys = Object.keys(months).sort();
    var row = findDate(monthsSheet, monthKeys[0]);
    if ( row == -1 ) { row = monthsSheet.getLastRow() + 1; } else {
      monthsSheet.getRange(row, 1, monthsSheet.getLastRow()-row+1, monthsSheet.getLastColumn()).clear();
    }
    row -= 1;
    for (const month of monthKeys) {
      row += 1;
      monthsSheet.getRange(row,1).setValue(month)
      for (const [key, value] of Object.entries(months[month])) {
        let index = headers.indexOf(key);
        if ( index == -1 ) {
          if (LMdebug) {Logger.log('didn\'t find %s', key);}
          index = headers.push(key) - 1; //push returns length, we want index which is one less
          monthsSheet.getRange(1, monthsSheet.getLastColumn()+1).setValue(key);
          monthsSheet.getRange(row,index+1).setValue(value.toFixed(2));
        } else {
          monthsSheet.getRange(row,index+1).setValue(value.toFixed(2));
        }
      }
    }
  }

  if (LMCoalesceDays) {
    var daysSheet = LMActiveSpreadsheet.getSheetByName('LM-Days');
    if (daysSheet == null) { daysSheet = createSheet('LM-Days'); }
    var headers = daysSheet.getRange(1, 1, 1, daysSheet.getLastColumn()).getValues()[0];
    var daysKeys = Object.keys(days).sort();
    var row = findDate(daysSheet, daysKeys[0]);
    if ( row == -1 ) { row = daysSheet.getLastRow() + 1; } else {
      daysSheet.getRange(row, 1, daysSheet.getLastRow()-row+1, daysSheet.getLastColumn()).clear();
    }
    row -= 1;
    for (const day of daysKeys) {
      row += 1;
      daysSheet.getRange(row,1).setValue(day)
      for (const [key, value] of Object.entries(days[day])) {
        let index = headers.indexOf(key);
        if ( index == -1 ) {
          if (LMdebug) {Logger.log('didn\'t find %s', key);}
          index = headers.push(key) - 1; //push returns length, we want index which is one less
          daysSheet.getRange(1, daysSheet.getLastColumn()+1).setValue(key);
          daysSheet.getRange(row,index+1).setValue(value.toFixed(2));
        } else {
          daysSheet.getRange(row,index+1).setValue(value.toFixed(2));
        }
      }
    }
  }
}

function findDate(sheet, date) {
  if (sheet.getName() == 'LM-Accounts') {
    var dates = sheet.getRange(3, 1, sheet.getLastRow()).getValues();
  } else {
    var dates = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  }
  let row = dates.findIndex(foo => {return foo[0] >= date});
  // if (LMdebug) {Logger.log('findDate row %s', row);}
  if (row == -1) { return -1; }
  if (sheet.getName() == 'LM-Accounts') { return row+3; }
  return row+2;
}

function createSheet(name) {
  var sheet = LMActiveSpreadsheet.insertSheet(name);
  if (name == 'LM-Accounts') {
    let headers = ['Date', 'Net Worth']
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, 1).setValue('Latest');
    let firstRow = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    firstRow.setNumberFormat("$#,##0.00;$(#,##0.00)");
    let secondRow = sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    secondRow.setNumberFormat("@");
    let dateCol = sheet.getRange("A1:A");
    dateCol.setNumberFormat("@");
    let theRest = sheet.getRange(3, 2, sheet.getMaxRows(), sheet.getMaxColumns());
    theRest.setNumberFormat("$#,##0.00;$(#,##0.00)");
    return sheet
  }
  var LMCategories = JSON.parse(LMDocumentProperties.getProperty('LMCategories'));
  LMCategories.sort((a, b) => a.order - b.order);
  var headers = [];
  headers.push('Date', 'Total Net', 'Total Exp');

  for (const category of LMCategories) {
    if (!category.exclude_from_budget) { headers.push(category.name); }
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  let dateCol = sheet.getRange("A1:A");
  dateCol.setNumberFormat("@");
  let firstRow = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  firstRow.setNumberFormat("@");
  let theRest = sheet.getRange(2, 2, sheet.getMaxRows(), sheet.getMaxColumns());
  theRest.setNumberFormat("$#,##0.00;$(#,##0.00)");

  return sheet;
}

function findCat(LMCategories, catId) {
  for (const category of LMCategories) {
    if (category.id == catId){
      return category;
    }
  }
  displayToastAlert("Couldn't find a category, make sure you have updated them");
  throw new Error("Couldn't find a category " + catId);
}

function apiRequest(url) {
  if (LMdebug) {Logger.log('apiRequest for %s', url);}
  let ssId = LMActiveSpreadsheet.getId();
  try {
    var LMKey = PropertiesService.getUserProperties().getProperty('LMKey_'+ssId);
    if (typeof LMKey !== 'string' || LMKey == '') {
      displayToastAlert("Can't find API key, make sure you have saved it");
      return false;
    }
  } catch (err) {
    if (LMdebug) {Logger.log('Failed with error %s', err.message);}
    displayToastAlert("Can't find API key, make sure you have saved it");
    return false;
  }
  var headers = {
    'Authorization' : 'Bearer ' + LMKey
  };
  var options = {
    'method' : 'get',
    'headers' : headers
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (f) {
    displayToastAlert('API access failed');
    if (LMdebug) {Logger.log('URLfetch with error %s', f.message);}
    return false;
  }
}

function calculateRelativeDates() {
  var startDate = new Date();
  var endDate = new Date();
  startDate.setMonth(startDate.getMonth() - LMTransactionsLookbackMonths, 1);
  endDate.setDate(endDate.getDate() + 2);
  startDate = Utilities.formatDate(startDate, LMScriptTimezone, "yyyy-MM-dd");
  endDate = Utilities.formatDate(endDate, LMScriptTimezone, "yyyy-MM-dd");
  return {startDate, endDate}
}

function loadTransactions(startDate, endDate) {
  let url = 'https://dev.lunchmoney.app/v1/transactions?&is_group=false&debit_as_negative=true&pending=true&start_date=' + startDate + '&end_date=' + endDate
  let result = apiRequest(url);
  try{
    if (result != false) {
      if (result.transactions.length > 0) {
        return result.transactions;
      } else {
        if (LMdebug) {Logger.log('loadTransactions empty for %s', url);}
        displayToastAlert("transactions seem empty for those dates");
        return false
      }
    } else {
      return false;
    }
  } catch (err) {
    if (LMdebug) {Logger.log('result %s', result);}
    if (LMdebug) {Logger.log('loadTransactions failed with error %s', err.message);}
    displayToastAlert("unable to retrieve transactions");
    return false
  }
}

function coalesce(object, period, parsedTransaction, name) {
  if (object.hasOwnProperty(period)) {
    if (object[period].hasOwnProperty(name)) {
      object[period][name] += parsedTransaction.to_base;
    } else {
      object[period][name] = parsedTransaction.to_base;
    }
  } else {
    object[period] = new Object();
    object[period][name] = parsedTransaction.to_base;
  }
}

function parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames) {
  var parsedTransactions_2d = [];
  var months = {};
  var days = {};

  for (const transaction of transactions) {
    let parsedTransaction = parseTransaction(transaction, LMCategories, plaidAccountNames, assetAccountNames);

    if (LMCoalesce && !parsedTransaction.exclude_from_budget) {
      let month = parsedTransaction.date.slice(0, -3);
      let day = parsedTransaction.date;
      coalesce(months, month, parsedTransaction, parsedTransaction.category_name);
      coalesce(days, day, parsedTransaction, parsedTransaction.category_name);
      if (parsedTransaction.hasOwnProperty('group_name')) {
        coalesce(months, month, parsedTransaction, parsedTransaction.group_name);
        coalesce(days, day, parsedTransaction, parsedTransaction.group_name);
      }
      if (!parsedTransaction.exclude_from_totals && !parsedTransaction.is_income) {
        coalesce(months, month, parsedTransaction, 'Total Exp');
        coalesce(days, day, parsedTransaction, 'Total Exp');
      }
      if (!parsedTransaction.exclude_from_totals) {
        coalesce(months, month, parsedTransaction, 'Total Net');
        coalesce(days, day, parsedTransaction, 'Total Net');
      }
      if (transaction.hasOwnProperty('tags') && transaction.tags != null && transaction.tags.length > 0) {
        for (const tag of transaction.tags) {
          coalesce(months, month, parsedTransaction, tag.name);
          coalesce(days, day, parsedTransaction, tag.name);
        }
      }
    }

    parsedTransactions_2d.push([parseInt(parsedTransaction.id), parsedTransaction.date, parsedTransaction.category_string, parsedTransaction.payee, parsedTransaction.to_base, parsedTransaction.notes, parsedTransaction.account_name, parsedTransaction.tag_string, parsedTransaction.status, parsedTransaction.exclude_from_totals, parsedTransaction.exclude_from_budget, parsedTransaction.is_income]);
  }
  return {parsedTransactions_2d, months, days}
}

function parseTransaction(transaction, LMCategories, plaidAccountNames, assetAccountNames) {
  //tags
  transaction.tag_string = '';
  if (transaction.hasOwnProperty('tags') && transaction.tags != null && transaction.tags.length > 0) {
    for (const tag of transaction.tags) {
      transaction.tag_string += tag.name;
      transaction.tag_string += ', ';
    }
    transaction.tag_string = transaction.tag_string.slice(0, -2); 
  }

  //categories
  if (transaction.category_id != null) {
    var category = findCat(LMCategories, transaction.category_id);
    transaction.category_name = category.name;
    if (category.group_id != null) {
      let groupCategory = findCat(LMCategories, category.group_id);
      transaction.category_string = groupCategory.name + '/' + transaction.category_name;
      transaction.group_name = groupCategory.name
    } else {
      transaction.category_string = transaction.category_name;
    }
    transaction.exclude_from_totals = category.exclude_from_totals;
    transaction.is_income = category.is_income;
    transaction.exclude_from_budget = category.exclude_from_budget;
  } else {
    transaction.category_string = '';
    transaction.exclude_from_totals = '';
    transaction.is_income = '';
    transaction.exclude_from_budget = ''; 
  }

  //account name
  if (transaction.hasOwnProperty('asset_id') && transaction.asset_id !== null) {
    if (!assetAccountNames.hasOwnProperty(transaction.asset_id)) {
      assetAccountNames = updateAssetAccountNames();
    }
    transaction.account_name = assetAccountNames[transaction.asset_id];
  } else if (transaction.hasOwnProperty('plaid_account_id') && transaction.plaid_account_id !== null) {
    if (!plaidAccountNames.hasOwnProperty(transaction.plaid_account_id)) {
      plaidAccountNames = updatePlaidAccountNames();
    }
    transaction.account_name = plaidAccountNames[transaction.plaid_account_id];
  } else {
    transaction.account_name = 'Cash'
  }

  //should to_base be negative
  if (+(transaction.amount) < 0) {
    transaction.to_base = 0 - transaction.to_base;
  } else {
    transaction.to_base = Math.abs(transaction.to_base);
  }

  return transaction
}

function loadCategoriesAndAccounts() {
  if (LMTrackPlaidAccounts) {
    var {plaidAccountNames, plaidAccounts} = updatePlaidAccounts();
  } else {
    try {
      var plaidAccountNames = JSON.parse(LMDocumentProperties.getProperty('LMPlaidAccountNames'));
      var plaidAccounts = null;
      if (plaidAccountNames == null) {
        var {plaidAccountNames, plaidAccounts} = updatePlaidAccounts()
        if (plaidAccountNames == false) {
          throw new Error('failed to load LMPlaidAccountNames 1');
        }
      }
    } catch (err) {
      if (LMdebug) {Logger.log('plaidAccountNames failed with error %s', err.message);}
      throw new Error('failed to load LMPlaidAccountNames 2');
    }
  }

  if (LMTrackAssets) {
    var {assetAccountNames, assetAccounts} = updateAssetAccounts();
  } else {
    try {
      var assetAccountNames = JSON.parse(LMDocumentProperties.getProperty('LMAssetAccountNames'));
      var assetAccounts = null;
      if (assetAccountNames == null) {
        var {assetAccountNames, assetAccounts} = updateAssetAccounts();
        if (assetAccountNames == false) {
          throw new Error('failed to load LMAssetAccountNames 1');
        }
      }
    } catch (err) {
      if (LMdebug) {Logger.log('assetAccountNames failed with error %s', err.message);}
      throw new Error('failed to load LMAssetAccountNames 2');
    }
  }
  
  try {
    var LMCategories = JSON.parse(LMDocumentProperties.getProperty('LMCategories'));
    if (typeof LMCategories !== 'object' || LMCategories == null) {
      LMCategories = updateCatagories();
    }
  } catch (err) {
    if (LMdebug) {Logger.log('categories failed with error %s', err.message);}
    return false;
  }

  return {LMCategories, plaidAccountNames, assetAccountNames, plaidAccounts, assetAccounts};

}

function updatePlaidAccounts() {
  let url = 'https://dev.lunchmoney.app/v1/plaid_accounts'
  let result = apiRequest(url);
  if (result != false) {
    var plaidAccounts = result.plaid_accounts;
  } else { throw new Error('api apiRequest failed in updatePlaidAccounts'); }

  var plaidAccountNames = {};
  for (const plaidAccount of plaidAccounts) {
    if (plaidAccount.display_name == '' || plaidAccount.display_name == null) {
      plaidAccount.display_name = htmlDecode(plaidAccount.name);
    }
    plaidAccountNames[plaidAccount.id] = htmlDecode(plaidAccount.display_name);
  }

  if (!LMTrackAssets) {
    LMDocumentProperties.setProperty('LMPlaidAccountNames', JSON.stringify(plaidAccountNames));
  }

  return {plaidAccountNames, plaidAccounts};
}

function updateAssetAccounts() {
  let url = 'https://dev.lunchmoney.app/v1/assets'
  let result = apiRequest(url);
  if (result != false) {
    var assetAccounts = result.assets;
  } else { throw new Error('api apiRequest failed in updatePlaidAccounts'); }

  var assetAccountNames = {};
  for (const asset of assetAccounts) {
    if (asset.display_name == '' || asset.display_name == null) {
      asset.display_name = htmlDecode(asset.name);
    }
    assetAccountNames[asset.id] = htmlDecode(asset.display_name);
  }

  if (!LMTrackAssets) {
    LMDocumentProperties.setProperty('LMAssetAccountNames', JSON.stringify(assetAccountNames));
  }

  return {assetAccountNames, assetAccounts};
}

function htmlDecode(input) {
  let decode = XmlService.parse('<d>' + input + '</d>');
  return decode.getRootElement().getText();
}

function updateCatagories() {
  let url = 'https://dev.lunchmoney.app/v1/categories'
  let result = apiRequest(url);
  if (result != false) {
    var categories = result.categories;
  } else {
    return
  }
  try {
    LMDocumentProperties.setProperty('LMCategories', JSON.stringify(categories));
    return categories
  } catch (err) {
    if (LMdebug) {Logger.log('Failed with error %s', err.message);}
    displayToastAlert("Can't save categories for some reason");
    return
  }
}

function displayPrompt(promptString) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(promptString);
  return result
}

function setApiKey() {
  let ssId = LMActiveSpreadsheet.getId();
  PropertiesService.getUserProperties().setProperty('LMKey_'+ssId, displayPrompt("enter API key").getResponseText());
  displayToastAlert("saved key")
}

function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert"); 
}

/**
 * Return total for a category or tag over months.
 *
 * @param {string} category The category or tag.
 * @param {string} startDate The start date in YYYY-MM.
 * @param {string} endDate The end date in YYYY-MM, if in future, will total until last row.
 * @param {string} random Some random to evade caching. Use 'LM-Transactions'!R1
 * @return The input multiplied by 2.
 * @customfunction
 */
function LMCATTOTAL(category, startDate, endDate, random) {
  if (LMdebug) {Logger.log('in LMCATTOTAL cat %s, start %s, end %s, random %s', category, startDate, endDate, random);}
  var monthsSheet = LMActiveSpreadsheet.getSheetByName('LM-Months');
  if (monthsSheet == null) { throw new Error('Need LM-Months'); }
  var headers = monthsSheet.getRange(1, 1, 1, monthsSheet.getLastColumn()).getValues()[0];
  let index = headers.indexOf(category);
  if ( index == -1 ) { throw new Error('Can\'t find ' + category); }
  var startRow = findDate(monthsSheet, startDate);
  if ( startRow == -1 ) { throw new Error('Can\'t find ' + startDate); }
  var endRow = findDate(monthsSheet, endDate);
  if ( endRow == -1 ) { endRow = monthsSheet.getLastRow(); }
  let foo = monthsSheet.getRange(startRow, index+1, endRow - startRow + 1, 1).getValues();
  var sum = 0;
  for (const val of foo) {
    sum += +(val);
  }
  if (LMdebug) {Logger.log('in LMCATTOTAL sum %s', sum);}
  return sum
}
