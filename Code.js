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
const LMdebug = true;         //write tracing info to the apps script log

const LMJumpOnFinish = true;    //should we jump to the last row when we finish updating a sheet?

const LMTransactionsLookbackDays = 60 //number of days back from the current last one, to check for updated
                                      //category, etc. This one you should keep tight if you can, it's a bit slow.

var LMTransactionsLookback = 1000     //Max number of transactions you would ever get from today to LMTransactionsLookbackDays.
                                      //Be generous, it's fast. If you start getting sets of old transactions added at the end of
                                      //your sheet, this is where the problem is, make it bigger.

var LMPendingLookback = 300;  //number of transactions back from the current last one, to check 
                              //for pending transactions that need to be updated. Be generous, it's fast.
/**
 *  END OF SETTINGS
 */

const LMDocumentProperties = PropertiesService.getDocumentProperties();
const LMActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var LMTransactionAllIds = null;
var LMTransactionAllIdsStart = null;

function updateTransactionsAll() {
  LMTransactionAllIds = null;
  LMTransactionAllIdsStart = null;
  var transactionsAllSheet = LMActiveSpreadsheet.getSheetByName("LM-Transactions-All");
  if (transactionsAllSheet == null) {
    let firstTransactionDate = displayPrompt("date of first transaction in yyyy-MM-dd format").getResponseText();
    var today = new Date();
    today = Utilities.formatDate(today, "GMT", "yyyy-MM-dd");
    let transactions = loadTransactions(firstTransactionDate, today);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let {LMCategories, plaidAccountNames, assetAccountNames} = loadCategoriesAndAccounts();
    let parsedTransactions = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
    transactionsAllSheet = createTransactionsAllSheet();
    transactionsAllSheet.getRange(2, 1, parsedTransactions.length, parsedTransactions[0].length).setValues(parsedTransactions);
  } else {
    var transactionsAllLastRow = transactionsAllSheet.getLastRow();
    let {LMCategories, plaidAccountNames, assetAccountNames} = loadCategoriesAndAccounts();
    checkPendings(LMCategories, plaidAccountNames, assetAccountNames, transactionsAllSheet, transactionsAllLastRow);
    let {startDate, endDate} = calulateRelitiveDates(LMTransactionsLookbackDays, transactionsAllSheet, transactionsAllLastRow);
    let transactions = loadTransactions(startDate, endDate);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let parsedTransactions = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
    for (const transaction of parsedTransactions) {
      let row = findIdTransactionsAll(transaction[0].toFixed(0), transactionsAllSheet, transactionsAllLastRow);
      if (row > 0) {
         transactionsAllSheet.getRange(row, 1, 1, transaction.length).setValues([transaction]);
      } else {
        transactionsAllLastRow = transactionsAllLastRow + 1;
        transactionsAllSheet.getRange(transactionsAllLastRow, 1, 1, transaction.length).setValues([transaction]);
      }
    }
  }
  if (LMJumpOnFinish) {transactionsAllSheet.setActiveCell(transactionsAllSheet.getDataRange().offset(transactionsAllLastRow, 0, 1, 1));}
}

function checkPendings(LMCategories, plaidAccountNames, assetAccountNames, transactionsAllSheet, transactionsAllLastRow) {
  var start = transactionsAllLastRow - LMPendingLookback;
  if (start < 1) {
    start = 1;
    LMPendingLookback = transactionsAllLastRow + 1;
  }
  let search = 'pending';
  var data = transactionsAllSheet.getRange(start, 9, LMPendingLookback).getValues();
  var ids = transactionsAllSheet.getRange(start, 1, LMPendingLookback).getValues();
  while (true) {
    let index = data.findIndex(foo => {return foo[0] == search});
    if (index < 0) {break;}
    let url = 'https://dev.lunchmoney.app/v1/transactions/' + ids[index] + '?debit_as_negative=true';
    let transaction = apiRequest(url);
    let parsedTransaction = parseTransaction(transaction, LMCategories, plaidAccountNames, assetAccountNames);
    let trasactionArray = [parseInt(parsedTransaction.id), parsedTransaction.date, parsedTransaction.category_name, parsedTransaction.payee, parsedTransaction.to_base, parsedTransaction.notes, parsedTransaction.account_name, parsedTransaction.tag_string, parsedTransaction.status, parsedTransaction.exclude_from_totals, parsedTransaction.exclude_from_budget, parsedTransaction.is_income];
    transactionsAllSheet.getRange(index + start, 1, 1, trasactionArray.length).setValues([trasactionArray]);
    data[index] = 'foo';
  }
}

function findIdTransactionsAll(id, transactionsAllSheet, transactionsAllLastRow){
  if (LMTransactionAllIds == null) {
    LMTransactionAllIdsStart = transactionsAllLastRow - LMTransactionsLookback;
    if (LMTransactionAllIdsStart < 1) {
      LMTransactionAllIdsStart = 1;
      LMTransactionsLookback = transactionsAllLastRow + 1;
    }
    LMTransactionAllIds = transactionsAllSheet.getRange(LMTransactionAllIdsStart, 1, LMTransactionsLookback).getValues();
  }

  let row = LMTransactionAllIds.findIndex(foo => {return foo[0] == id});
  return row + LMTransactionAllIdsStart
}

function createTransactionsAllSheet() {
  var transactionsAllSheet = LMActiveSpreadsheet.insertSheet('LM-Transactions-All');
  let data = [['id', 'date', 'category name', 'payee', 'amount', 'notes', 'account name', 'tag', 'status', 'exclude from totals', 'exclude from budget', 'is income']]
  transactionsAllSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return transactionsAllSheet
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

function calulateRelitiveDates(lookback, transactionsAllSheet, transactionsAllLastRow) {
  let lastDate = transactionsAllSheet.getRange(transactionsAllLastRow, 2, 1).getValue();
  var startDate = new Date(lastDate);
  var endDate = new Date();
  startDate.setDate(startDate.getDate() - lookback);
  endDate.setDate(endDate.getDate() + 2);
  startDate = Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd");
  endDate = Utilities.formatDate(endDate, "GMT", "yyyy-MM-dd");
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

function parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames) {
  var transactions_2d = [];

  for (const transaction of transactions) {
    let parsedTransaction = parseTransaction(transaction, LMCategories, plaidAccountNames, assetAccountNames);
    transactions_2d.push([parseInt(parsedTransaction.id), parsedTransaction.date, parsedTransaction.category_name, parsedTransaction.payee, parsedTransaction.to_base, parsedTransaction.notes, parsedTransaction.account_name, parsedTransaction.tag_string, parsedTransaction.status, parsedTransaction.exclude_from_totals, parsedTransaction.exclude_from_budget, parsedTransaction.is_income]);
  }
  return transactions_2d
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
      category = findCat(LMCategories, category.group_id);
      transaction.category_name = category.name + '/' + transaction.category_name;
    }
    transaction.exclude_from_totals = category.exclude_from_totals;
    transaction.is_income = category.is_income;
    transaction.exclude_from_budget = category.exclude_from_budget;
  } else {
    transaction.category_name = '';
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
  try {
    var plaidAccountNames = JSON.parse(LMDocumentProperties.getProperty('LMPlaidAccountNames'));
    if (plaidAccountNames == null) {
      plaidAccountNames = updatePlaidAccountNames()
      if (plaidAccountNames == false) {
        throw new Error('failed to load LMPlaidAccountNames 1');
      }
    }
  } catch (err) {
    if (LMdebug) {Logger.log('plaidAccountNames failed with error %s', err.message);}
    throw new Error('failed to load LMPlaidAccountNames 2');
  }

  try {
    var assetAccountNames = JSON.parse(LMDocumentProperties.getProperty('LMAssetAccountNames'));
    if (assetAccountNames == null) {
      assetAccountNames = updateAssetAccountNames();
      if (assetAccountNames == false) {
        throw new Error('failed to load LMAssetAccountNames 1');
      }
    }
  } catch (err) {
    if (LMdebug) {Logger.log('assetAccountNames failed with error %s', err.message);}
    throw new Error('failed to load LMAssetAccountNames 2');
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

  return {LMCategories, plaidAccountNames, assetAccountNames}

}

function updatePlaidAccountNames() {
  let url = 'https://dev.lunchmoney.app/v1/plaid_accounts'
  let result = apiRequest(url);
  if (result != false) {
    var plaidAccounts = result.plaid_accounts;
  } else {
    return false
  }
  var plaidAccountNames = {};
  for (const plaidAccount of plaidAccounts) {
    if (plaidAccount.display_name == '' || plaidAccount.display_name == null) {
      plaidAccount.display_name = plaidAccount.name;
    }
    plaidAccountNames[plaidAccount.id] = plaidAccount.display_name;
  }
  try {
    LMDocumentProperties.setProperty('LMPlaidAccountNames', JSON.stringify(plaidAccountNames));
  } catch (err) {
    if (LMdebug) {Logger.log('Saving LMPlaidAccountNames failed with error %s', err.message);}
    return false
  }
  return plaidAccountNames
}

function updateAssetAccountNames() {
  let url = 'https://dev.lunchmoney.app/v1/assets'
  let result = apiRequest(url);
  if (result != false) {
    var assets = result.assets;
  } else {
    return false
  }
  var assetAccountNames = {};
  for (const asset of assets) {
    if (asset.display_name == '' || asset.display_name == null) {
      asset.display_name = asset.name;
    }
    assetAccountNames[asset.id] = asset.display_name;
  }
  try {
    LMDocumentProperties.setProperty('LMAssetAccountNames', JSON.stringify(assetAccountNames));
  } catch (err) {
    if (LMdebug) {Logger.log('Saving LMAssetAccountNames failed with error %s', err.message);}
    return false
  }
  return assetAccountNames
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