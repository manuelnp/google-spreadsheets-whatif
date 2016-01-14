/**
 * Replaces whatif_param with whatif_value anywhere in formula_addr hieranchy
 *
 * @param {string} formula_addr String with formula cell address, not the actual cell reference!
 * @param {string} whatif_param String with whatif param cell address, not the actual cell reference!
 * @return {string} whatif_values String with whatif values range address, not the actual range reference!
 */
function WhatIf(formula_addr, whatif_param, whatif_values) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  //var values = sheet.getRange(whatif_values).getValues();
  
  var resolved = {}
  resolved[whatif_param] = whatif_values;
  
  return Resolve_R(formula_addr, whatif_param, whatif_values, resolved, sheet);
}

function Resolve_R(formula_addr, whatif_param, whatif_values, resolved, sheet) {
  var formulas = []
  var formula = sheet.getRange(formula_addr).getFormula();
  if (formula != "") {
    for (var i = 0; i < whatif_values.length; i++) {
      formulas[i] = formula;
    }
    var params = formula.match(/[A-Z]+[0-9]+/g);
    for (var i = 0; i < params.length; i++) {
      var param = params[i];
      if (!(param in resolved)) {
        resolved[param] = Resolve_R(param, whatif_param, whatif_values, resolved, sheet);
      }
      for (var j = 0; j < whatif_values.length; j++) {
        formulas[j] = formulas[j].replace(new RegExp(param, 'g'), resolved[param][j]);
      }
    }
    for (var i = 0; i < whatif_values.length; i++) {
      formulas[i] = formulas[i].substring(1); // Remove leading '='
      formulas[i] = eval(formulas[i]);
    }
  } else {
    for (var i = 0; i < whatif_values.length; i++) {
      var value = sheet.getRange(formula_addr).getValue();
      formulas[i] = value;
    }
  }

  return formulas;
}
