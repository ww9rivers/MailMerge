/**
 * Common utilities.
 *
 * This code is licensed under the GPL v3.0, which is found at the URL below:
 *     http://opensource.org/licenses/gpl-3.0.html
 *
 * Copyright (c) 2014 9Rivers.com. All rights reserved.
 *
 * To do:
 *    - Make the functions configurable with formats.
 */
function CommonUtils() {}
CommonUtils.prototype.address = function(addr) {
  return addr ? addr.replace(/,\s*/g, "\n") : "";
};
CommonUtils.prototype.date = function(date) {
  return Utilities.formatDate(date, 'EST', "yyyy-MM-dd");
};
CommonUtils.prototype.money = function(amount) {
  return amount.toFixed(2);
};
CommonUtils.prototype.time = function(date) {
  return Utilities.formatDate(date, 'EST', "yyyy-MM-dd HH:mm:ss");
};
CommonUtils.prototype.zint = function(nmbr) {
  var padding = 4;
  return (new Array(padding).join('0')+nmbr).slice(-padding);
};

function test() {
  co = new CommonUtils;
  Logger.log("Result = ["+co.address('a,b, c')+']');
  Logger.log("Current time is: "+co.time(new Date));
};
