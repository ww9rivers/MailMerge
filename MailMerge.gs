/**
 * A mail merge program using Google Docs, Sheets and Apps Script services.
 *
 * This is created for the Ann-Hua Rainbow Project (http://rainbow.annhua.org/), a charity
 * program supporting underprivileged young Chinese students for their primary and secondary
 * level education. Your support would be very much appreciated.
 *
 * This code is licensed under the GPL v3.0, which is found at the URL below:
 *     http://opensource.org/licenses/gpl-3.0.html
 *
 * Copyright (c) 2014 9Rivers.com. All rights reserved.
 */

var c9r = {
  /** The current spreadsheet */
  doc: null,

  /** Google Docs URI */
  GDocs: "https://docs.google.com/document/d/",
  /** Reference manual */
  Ref: "19ULlrVs6ykWjsPZ27mSoev4PetrsjA97WNss4wTgoI4/pub",
  /** Default settings -- All may be set in the "Settings" sheet. */
  defaultSetting: {
    list: 'Donations',
    format_Date: 'yyyy-MM-dd',
    sender: "Rainbow@annhua.org",
    today: new Date(),
    OnBehalf: '',	// {OnBehalf} in message body
    Year: 2016,		// {Year} in message body
    levels: 0,		// Sponsor levels
    spondor: '',	// Sponsor message
    timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone()
  },
  /** MasterSheet: initialized by onOpen(). */
  ms: false,
  /** Seetings loaded from the "Settings" sheet, combined with defaultSettings */
  setting: false,
  /** The "list" sheet */
  listsheet: null,
  /** Dictionary from column name to number. */
  cols: {
    Address:	13,
    Amount:		1,
    Date:		4,
    Name:		2,
    No:			0,
    OnBehalf:	3
  },
  receiptCol: 'H',
  
  /** Convert specified row from a list to a dictionary.
   *
   * @param {list} vals  A list of values, representing a row in the list sheet.
   * @return {Array} of the row converted from list to a dictionary.
   */
  row: function(vals) {
    var dx = {};
    for (var k in c9r.cols) {
      dx[k] = vals[c9r.cols[k]];
    }
    return dx;
  },

  /** Get the current spreadsheet object.
   */
  init: function() {
    if (c9r.doc == null) {
      c9r.util = new CommonUtils;
      let doc = c9r.doc = SpreadsheetApp.getActiveSpreadsheet();
      c9r.defaultSetting.sender = (doc.getOwner()||doc.getEditors()[0]).getEmail();
      c9r.defaultSetting.Year = c9r.defaultSetting.today.getFullYear()-1;
      var settings = c9r.setting = getDict('Settings', c9r.defaultSetting);
      var listsheet = c9r.listsheet = doc.getSheetByName(settings.list);
      settings.subject = settings.subject.replace('{Year}', settings.Year);
      settings.body = settings.body.replace('{Year}', settings.Year);
      // Set up c9r.cols using the settings values:
      var hd2col = {};
      var ncols = listsheet.getMaxColumns();
      Logger.log("Sheet "+settings.list+' has '+ncols+' columns');
      var row1 = listsheet.getRange(1, 1, 1, ncols).getValues()[0];
      Logger.log("Row 1: "+row1.toString());
      for (var k in row1) {
        hd2col[row1[k]] = k;
        Logger.log("hd2col: ["+row1[k]+'] =>'+k);
      }
      for (var k in settings) {
        var val = settings[k];
        if (val[0] == '[') {
          var key = (val != '[]' ? val : ('['+k+']'));
          c9r.cols[k] = hd2col[key];
        }
      }
      c9r.ReceiptCol = String.fromCharCode(65+parseInt(c9r.cols.Receipt));
      Logger.log("Receipt column: "+c9r.ReceiptCol+' : '+c9r.cols.Receipt);
    }
    return c9r.setting;
  },
  
  /** Set active cell to "[Receipt]" column row /i/.
   *
   * @param {integer} i  Row number in the [Receipt] column.
   */
  receipt: function(i) {
    var rcell = c9r.ReceiptCol+(i+1);
    Logger.log("Going to cell: "+rcell+' for row '+i);
    return c9r.listsheet.setActiveCell(rcell);
  }
};

/**
 * Traverse through rows in the sheet, execute function /run/ with the
 * given options in /opt/.
 *
 * @param {function} run  Function to run on each row.
 * @param {object} opt    Options for this and /run/ functions.
 */
function forEachRow(run, opt) {
  var sheet = c9r.listsheet;
  var range = sheet.getDataRange();
  var rows = range.getNumRows();
  var vals = range.getValues();
  for (var i = 0; i < rows; i++) {
    if (run(vals[i], i, opt) === false) {
      break;
    }
  }
};

/**
 * Read a sheet with given /ssname/ and return the contents as a dictionary,
 * with column A as key and B as value.
 *
 * If a key is enclosed in square brackets, as [key], that row is ignored.
 *
 * @param {string} ssname   Name of a sheet in the current document.
 * @param {Array} defaults  Default values for the keys in the dictionary if not specified in the sheet.
 * @return A dictionary of { key: value } from the named sheet.
 */
function getDict(ssname, defaults) {
  var sheet = c9r.doc.getSheetByName(ssname);
  var range = sheet.getDataRange();
  var rows = range.getNumRows();
  var vals = range.getValues();
  var dict = {}
  for (var i = 0; i < rows; i++) {
    var key = vals[i][0];
    if (key[0] != '[') {
      dict[key] = vals[i][1];
    }
  }
  if (defaults) {
    for (var k in defaults) {
      if (!(k in dict)) {
        dict[k] = defaults[k];
      }
    }
  }
  return dict;
};

/**
 * Open a pop-up to show a link to the reference manual.
 */
function helpMailMerge() {
  var app = UiApp.createApplication().setHeight(50).setWidth(200);
  app.setTitle("Mail Merge Project Help");
  var link = app.createAnchor('Open in new window', c9r.GDocs+c9r.Ref).setId("link");
  app.add(link);  
  SpreadsheetApp.getActive().show(app);
};

/**
 * Handler function for menu item to merge documents.
 */
function mergeMail() {
  if (Browser.msgBox(
      'Please confirm',
      'Are you ready to create receipt documents?',
      Browser.Buttons.YES_NO) != 'yes') {
    return;
  }

  var setting = c9r.init();
  var temp = DriveApp.getFileById(setting.template);
  var opt = {
    setting: setting,
    temp: temp,
    folder: DriveApp.getFolderById(setting.folder),
    tempname: temp.getName()
  };
  // c9r.util.date()
  c9r.doc.toast('Today is '+(opt.setting.today.toString())
    +". Template is "+opt.tempname
    +'. Receipt folder is: '+opt.folder.getName());
  forEachRow(mergeRow, opt);
};

/** Merge one row.
 *
 * @param {Array} row   One row in the "Donations" sheet;
 * @param {integer} i   Row number;
 * @param {object} opt  Options
 */
function mergeRow(row, i, opt) {
  var rod = c9r.row(row); // The "row" as dictionary
  var nmbr = parseInt(rod.No, 10);
  var amount = parseFloat(rod.Amount);
  if (nmbr <= 0 || isNaN(nmbr) || isNaN(amount) || amount <= 0 || rod.Receipt != '') {
    return;
  }
  // To-Do: Need to get the column numbers by name.
  var name = rod.Name;
  var student = rod.Student;
  var date = rod.Date;
  var mailto = rod.Email;
  var addr = c9r.util.address(rod.Address);
  var recpt = c9r.util.zint(nmbr);
  var copyname = opt.tempname.replace('Template', '')+recpt+'-'+name;
  var copyFile = opt.temp.makeCopy(copyname, opt.folder);
  var copyKey = copyFile.getId();
  var copyDoc = DocumentApp.openById(copyKey);
  var copyBody = copyDoc.getBody();
  copyBody.replaceText('{Year}', opt.setting.Year);
  copyBody.replaceText('{No}', recpt);
  copyBody.replaceText('{Address}', addr);
  copyBody.replaceText('{Name}', name);
  copyBody.replaceText('{Amount}', amount.toFixed(2));
  // Need to use: opt.setting.timezone, opt.setting.format_Date
  copyBody.replaceText('{Date}', c9r.util.date(date));
  copyBody.replaceText('{Today}', c9r.util.date(opt.setting.today));
  var onbehalf = opt.setting.OnBehalf;
  copyBody.replaceText('{OnBehalf}', (student && onbehalf) ? onbehalf.replace('{Student}', student) : "");
  // Figure out (total = Amount + Match) to determine sponsor level
  var sponsor = opt.setting.sponsor;
  var label = null;
  if (sponsor != '') {
    if ('Match' in rod) { amount += rod.Match; }
    for (var k = opt.setting.levels; k > 0; k--) {
      var level = 'level'+k;
      if (amount >= opt.setting[level]) {
        label = opt.setting['label'+k];
        if (label != '') label += ' ';
        sponsor = ' '+sponsor.replace('{label}', label);
        break;
      }
    }
  }
  copyBody.replaceText('{Sponsor}', label == null ? '' : sponsor);
  // moveFileTo(copyFile, opt.folder);
  copyDoc.saveAndClose();
  c9r.receipt(i).setFormula('=hyperlink("'+c9r.GDocs+copyKey+'", "'+recpt+'")');
  c9r.doc.toast("Letter created: "+copyname+" <"+mailto+">");
};

/** Menu function to go through all the rows and send out mail messages alredy created.
 */
function sendMergedMail() {
  if (Browser.msgBox(
      'Please confirm',
      'Are you ready to send out emails?',
      Browser.Buttons.YES_NO) != 'yes') {
    // User did not clicked "Yes".
    return;
  }

  var setting = c9r.init();
  var opt = {
    setting: setting,
    folder: DriveApp.getFolderById(setting.folder)
  };
  c9r.doc.toast('Sending receipts in folder: '+opt.folder.getName());
  forEachRow(sendRow, opt);
};

/** Send out email for one row.
 *
 * @param {Array} row   One row in the "Donations" sheet;
 * @param {integer} i   Row number;
 * @param {object} opt  Options
 */
function sendRow(row, i, opt) {
  var rod = c9r.row(row); // The "row" as dictionary
  if (rod.Receipt == '#') {
    return false; // end of run.
  }

  var nmbr = parseInt(rod.No, 10);
  var sendto = rod.Email;
  var cell = c9r.receipt(i);
  var sent = cell.getNote();
  if (sent.length > 0) {
    c9r.doc.toast('Receipt'+nmbr+' is already marked: '+sent+'.');
    return true;
  }
  var copyId = cell.getFormula().match(/=hyperlink\("([^"]+)",\s"\d+"\)/i);
  if (copyId) {
    copyId = copyId[1].split('/').pop();
  }
  if (isNaN(nmbr) || nmbr <= 0 || sendto == '' || copyId == null || copyId == '') {
    return true; // continue the run.
  }
  if (opt.setting.interactive) {
    if (Browser.msgBox(
        'Please confirm',
        'Are you ready to send out receipt '+nmbr+' to '+rod.Name+' <'+sendto+'>',
        Browser.Buttons.YES_NO) != 'yes') {
      // User did not clicked "Yes".
      return Browser.msgBox(
        'Please confirm',
        'Do you want to continue?',
        Browser.Buttons.YES_NO) == 'yes';
    }
  }
  c9r.doc.toast('Sending receipt'+nmbr+' ('+copyId+') to <'+sendto+'>.');
  // Now send email:
  var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
  MailApp.sendEmail(sendto, opt.setting.subject, opt.setting.body,
    {
      replyTo: opt.setting.sender,
      htmlBody: opt.setting.body.replace(/\n/g, '<br/>'),
      attachments: pdf
    });
  cell.setNote(c9r.util.time(new Date));
  DriveApp.getFileById(copyId).setTrashed(true);
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  c9r.init();
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem("Create letters", "mergeMail")
    .addItem("Send email", "sendMergedMail")
    .addSeparator()
    .addItem("Help on Mail Merge", "helpMailMerge")
    .addToUi();
};