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

const LMJumpOnFinish = false;    //should we jump to the last row when we finish updating a sheet?

const LMTransactionsLookbackMonths = 1//number of full months we will pull transactions from, prior to the current one, to check
                                      //for updated category, etc. This one you should keep tight if you can, to reduce load on Lunch Money.

const LMTransactionsLookback = 1000   //Max number of transactions you would ever get from today to LMTransactionsLookbackMonths.
                                      //Be generous, it's fast. Script will error with a warning if it is too small.
/**
 *  END OF SETTINGS
 */

const LMDocumentProperties = PropertiesService.getDocumentProperties();
const LMActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

function updateTransactionsAll() {
  var transactionsAllSheet = LMActiveSpreadsheet.getSheetByName("LM-Transactions-All");
  if (transactionsAllSheet == null) {
    let firstTransactionDate = '1970-01-01';
    var today = new Date();
    today = Utilities.formatDate(today, "GMT", "yyyy-MM-dd");
    let transactions = loadTransactions(firstTransactionDate, today);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let {LMCategories, plaidAccountNames, assetAccountNames} = loadCategoriesAndAccounts();
    let {parsedTransactions_2d, parsedTransactions} = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
    transactionsAllSheet = createTransactionsAllSheet();
    transactionsAllSheet.getRange(2, 1, parsedTransactions_2d.length, parsedTransactions_2d[0].length).setValues(parsedTransactions_2d);
    var transactionsAllLastRow = transactionsAllSheet.getLastRow();
  } else {
    var transactionsAllLastRow = transactionsAllSheet.getLastRow();
    let {LMCategories, plaidAccountNames, assetAccountNames} = loadCategoriesAndAccounts();
    let {startDate, endDate} = calulateRelativeDates();
    let transactions = loadTransactions(startDate, endDate);
    if (transactions == false) {throw new Error("problem loading transactions");}
    let {parsedTransactions_2d, parsedTransactions} = parseTransactions(transactions, LMCategories, plaidAccountNames, assetAccountNames);
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
    // transactionsAllSheet.deleteRows(row+transactionsLength, 10); //not an off by one error, we want delete from the row after
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
  let LMTransactionAllIds = transactionsAllSheet.getRange(transactionAllIdsStart, 1, transactionsLookback).getValues();

  let row = LMTransactionAllIds.findIndex(foo => {return foo[0] == id});
  return row + transactionAllIdsStart
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

function calulateRelativeDates() {
  var startDate = new Date();
  var endDate = new Date();
  startDate.setMonth(startDate.getMonth() - LMTransactionsLookbackMonths, 1);
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

    if (!parsedTransaction.exclude_from_budget) {
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
        coalesce(months, month, parsedTransaction, 'Total');
        coalesce(days, day, parsedTransaction, 'Total');
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
  if (LMdebug) {Logger.log('months: %s', months);}
  if (LMdebug) {Logger.log('days: %s', days);}
  return {parsedTransactions_2d, months}
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