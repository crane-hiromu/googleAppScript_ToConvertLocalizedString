"use strict";

/*
  先を見据え一部ES6に沿って書いたが、letが対応していないため、宣言がconst,varのいずれかになっている。
*/

/* --------------------------------- Objects --------------------------------- */
/* ----------------- Resources　----------------- */
var RESOURCES = RESOURCES || {};

RESOURCES = {

  MESSAGELIST: 'MessageList', //メッセージ一覧のシート名。変更する場合はこちらも変更すること。
  MESSAGEID: 'message_id', //メッセージIDのカラム名。変更する場合はこちらも変更すること。

  LANGUAGES: {　
    EN: "en",
    JA: "ja"　
  },

  DIALOG_MESSAGE: {
    IOS: 'ios: ローカライズファイル',
    ANDROID: 'android: ローカライズファイル'
  },

  FILE_NAME: {
    IOS: 'Localizable.strings',
    ANDROID: 'strings.xml'
  },

  APPEND_LINES: {
    IOS: {
      START: [['/*'], ['  Localizable.strings'], ['*/']]
    },
    ANDROID: {
      START: [['<?xml version="1.0" encoding="utf-8"?> '], ['<resources>']],
      END: [['</resources>']]
    }
  },

  TXT_CONTENT_TYPE: 'text/comma-separated-values',
  TEXT_TYPE: 'utf-8',

  HTML: {
    MODAL: 'download_modal'
  },

  TYPE_ANDROID: {
    STRING: 's',
    INT: 'd'
  },
}

/* ----------------- Main Functions　----------------- */
var LOCALIZED_STRINGS = LOCALIZED_STRINGS || {};

LOCALIZED_STRINGS = {

  init: function() {
    this.controllers = this.CONTROLLERS;
    this.controllers.init();

    this.common = this.controllers.COMMON;
  },

  generate: function() {
    this.controllers.setLocalizedParameters();
    this.controllers.setUpLocalizedStringsSheets();
    this.controllers.writeLocalizedStringsSheets();
  },

  remove: function() {
    this.controllers.deleteStringsSheet(this.common.getStringsSheetName(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.EN));
    this.controllers.deleteStringsSheet(this.common.getStringsSheetName(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.JA));
    this.controllers.deleteStringsSheet(this.common.getStringsSheetName(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.EN));
    this.controllers.deleteStringsSheet(this.common.getStringsSheetName(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.JA));
  },

  downloadForIos: function() {
    const linkOfEnglish = this.controllers.getDownloadLink(this.common.getStringsSheetName(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.EN));
    const linkOfJapanese  = this.controllers.getDownloadLink(this.common.getStringsSheetName(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.JA));
    this.controllers.setDownloadLink(linkOfEnglish, linkOfJapanese, false, false, linkOfEnglish, RESOURCES.DIALOG_MESSAGE.IOS);
  },

  downloadForAndroid: function() {
    const linkOfEnglish = this.controllers.getDownloadLink(this.common.getStringsSheetName(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.EN));
    const linkOfJapanese  = this.controllers.getDownloadLink(this.common.getStringsSheetName(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.JA));
    this.controllers.setDownloadLink(false, false, linkOfEnglish, linkOfJapanese, linkOfEnglish, RESOURCES.DIALOG_MESSAGE.ANDROID);
  },

};


/* ----------------- Controller Functions 　----------------- */
LOCALIZED_STRINGS.CONTROLLERS = {

  init: function() {
    this.currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.common = this.COMMON;
    this.callback = this.CALLBACK;

    this.generate = this.GENERATE;
    this.generate.init(this.common, this.callback);

    this.download = this.DOWNLOAD;
    this.download.init(this.common);
  },

  setLocalizedParameters: function() {
    this.messageListSpreadSheet = this.common.getSheetRange(this.currentSheet.getSheetByName(RESOURCES.MESSAGELIST));
    this.spreadSheetSecondLine = this.messageListSpreadSheet[1];
    this.messageIdColumnNumber = this.spreadSheetSecondLine.indexOf(RESOURCES.MESSAGEID);
    this.startColumn = this.messageIdColumnNumber + 1;
    this.endColumn = this.startColumn + Object.keys(RESOURCES.LANGUAGES).length;
  },

  setUpLocalizedStringsSheets: function() {
    for (var i = this.startColumn; i < this.endColumn; i++) {
      var languageType = this.spreadSheetSecondLine[i];
      this.generate.generateStringsSheet(RESOURCES.FILE_NAME.IOS, languageType);
      this.generate.generateStringsSheet(RESOURCES.FILE_NAME.ANDROID, languageType);
    }
  },

  writeLocalizedStringsSheets: function() {
    this.generate.appendStartLines();

    const startLines = 2; //stringファイルに記載するテキスト部分が始まっている行番号のため固定。シートの上部目次ラインを変えない限りは値を変えないこと。
    const endLines = this.messageListSpreadSheet.length;
    for (var i = startLines; i < endLines; i++) {
      var messageLine = this.messageListSpreadSheet[i];
      this.generate.separateLines(messageLine);
    }

    this.generate.appendEndLines();
  },

  deleteStringsSheet: function(fileName) {
    const sheet = this.common.searchSheetByName(fileName);
    if(sheet) {
      this.currentSheet.deleteSheet(sheet);
    };
  },

  getDownloadLink: function(name) {
    this.common.resetSheet(name);

    const targetSheet = this.common.searchSheetByName(name);
    if(!targetSheet) { log("search sheet error"); return };

    const convertedSheet = this.download.convertSheet(this.common.getSheetRange(targetSheet));
    if(!convertedSheet) { log("convert error"); return };

    const downloadFileName = this.download.getFileNameForDownload(name);
    if(!downloadFileName) { log("replace error"); return };

    const driveFolder = DriveApp.getRootFolder().createFolder(name);
    const url = this.download.getDriveUrl(downloadFileName, convertedSheet, driveFolder);

    return url;
  },

  setDownloadLink: function(isIosEn, isIosJa, isAndroidEn, isAndroidJa, isModal, dialogMessage) {
    //以下4つはhtml側の決められたIDを参照しており、そこにURLをセットしている。そのためvarなどの宣言が不要。falseをセットすることでhtml側に出さないようにしている。
    linkForIosOfEnglish = isIosEn;
    linkForIosOfJapanese = isIosJa;
    linkForAndroidOfEnglish = isAndroidEn;
    linkForAndroidOfJapanese = isAndroidJa;

    const modalMessage = isModal ? dialogMessage : 'error: ローカライズファイルが生成されていません。'; //error handling
    this.download.displayDownloadModalView(modalMessage);
  },

}

/* ----------------- Common　Functions ----------------- */
LOCALIZED_STRINGS.CONTROLLERS.COMMON = {

  getSheetRange: function(sheet) {
    return sheet.getDataRange().getValues();
  },

  searchSheetByName: function(name) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  },

  getStringsSheetName: function(name, type) {
    return name + "(" + type + ")";
  },

  resetSheet: function(name) {
    const googleDriveFolder = DriveApp.getRootFolder().getFoldersByName(name);
    if(googleDriveFolder.hasNext()) {
      DriveApp.removeFolder(googleDriveFolder.next());
    };
  },

}

/* ----------------- Generate　Functions ----------------- */
LOCALIZED_STRINGS.CONTROLLERS.GENERATE = {

  init: function(common, callback) {
    this.currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.common = common;
    this.callback = callback;

    this.setLocalizedParameters();
  },

  setLocalizedParameters: function() {
    this.messageListSpreadSheet = this.common.getSheetRange(this.currentSheet.getSheetByName(RESOURCES.MESSAGELIST));
    this.spreadSheetSecondLine = this.messageListSpreadSheet[1];
    this.messageIdColumnNumber = this.spreadSheetSecondLine.indexOf(RESOURCES.MESSAGEID);
    this.startColumn = this.messageIdColumnNumber + 1;
    this.endColumn = this.startColumn + Object.keys(RESOURCES.LANGUAGES).length;
  },

  generateStringsSheet: function(name, type) {
    const fileName = this.common.getStringsSheetName(name, type);
    const sheet = this.common.searchSheetByName(fileName);
    sheet ? sheet.clear() : this.insertSheet(fileName);
  },

  insertSheet: function(name) {
    this.currentSheet.insertSheet(name);
  },

  getColumnNumber: function(languageType) {
    if(languageType == RESOURCES.LANGUAGES.EN) {
      return this.spreadSheetSecondLine.indexOf(RESOURCES.LANGUAGES.EN);

    } else if(languageType == RESOURCES.LANGUAGES.JA) {
      return this.spreadSheetSecondLine.indexOf(RESOURCES.LANGUAGES.JA);

    } else {
      log('error: non column number');
    }
  },

  separateLines: function(line) {
      const messageId = line[this.messageIdColumnNumber];

      for (var k = this.startColumn; k < this.endColumn; k++) {
        var language = this.spreadSheetSecondLine[k];
        var columnNumber = this.getColumnNumber(language);
        var message = line[columnNumber].replace(/\r?\n/g,"\\n");

        var messageForIos = this.getConvertMessage(messageId, message, this.callback.getReplacementStringForIos.bind(this), this.callback.getLineCharacterForIos);
        var messageForAndroid = this.getConvertMessage(messageId, message, this.callback.getReplacementStringForAndroid.bind(this), this.callback.getLineCharacterForAndroid);

        this.appendLocalizedStringLine(RESOURCES.FILE_NAME.IOS, language, messageForIos);
        this.appendLocalizedStringLine(RESOURCES.FILE_NAME.ANDROID, language, messageForAndroid);
      }
  },

  getConvertMessage: function(id, msg, getReplacementString, getLineCharacter) {
    var message = msg;
    if(message.match(/\{\{.*?\}\}/)) {
      for(var i=1; i <= message.split('{{').length; i++) { //置換されるindex番号を1からにしたいためにあえてi=1にしている。
        message = message.replace(/\{\{.*?\}\}/, getReplacementString(i, message));
      }
    }
    if(message.match(/\(\(.*?\)\)/)) {
      for(var k=0; k <= message.split('((').length; k++) {
        message = message.replace(/\(\(.*?\)\)/, getReplacementString(k, message));
      }
    }
    return getLineCharacter(id, message);
  },

  appendStartLines: function() {
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.EN, RESOURCES.APPEND_LINES.IOS.START);
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.IOS, RESOURCES.LANGUAGES.JA, RESOURCES.APPEND_LINES.IOS.START);
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.EN, RESOURCES.APPEND_LINES.ANDROID.START);
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.JA, RESOURCES.APPEND_LINES.ANDROID.START);
  },

  appendEndLines: function() {
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.EN, RESOURCES.APPEND_LINES.ANDROID.END);
    this.appendLinesOfSheet(RESOURCES.FILE_NAME.ANDROID, RESOURCES.LANGUAGES.JA, RESOURCES.APPEND_LINES.ANDROID.END);
  },

  appendLocalizedStringLine: function(name, type, stringLine) {
    const fileName = this.common.getStringsSheetName(name, type);
    const sheet = this.common.searchSheetByName(fileName);
    sheet.appendRow(stringLine);
  },

  appendLinesOfSheet: function(name, type, textArray){
    const sheet = this.common.searchSheetByName(this.common.getStringsSheetName(name, type));
    for (i in textArray) {
      sheet.appendRow(textArray[i]);
    }
  },

  getSubstr: function(str) {
    return str.substr(2).substr(0, str.length-4); //　最初と最後の括弧、4文字分を引いている。
  },

}

/* ----------------- Download　Functions ----------------- */
LOCALIZED_STRINGS.CONTROLLERS.DOWNLOAD = {

  init: function(common) {
    this.currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.common = common;
  },

  convertSheet: function(data){
    const rowlength = data.length;
    const columnlength = data[0].length;
    var convertedData = "";

    for(var i = 0; i < rowlength; i++){
      if(i < rowlength-1) {
        convertedData += data[i].join(",") + "\r\n";
      }else{
        convertedData += data[i];
      }
    }
    return convertedData;
  },

  getDriveUrl: function(fileName, sheet, driveFolder) {
    const blob = Utilities.newBlob("", RESOURCES.TXT_CONTENT_TYPE, fileName).setDataFromString(sheet, RESOURCES.TEXT_TYPE);
    const driveFile = DriveApp.getFolderById(driveFolder.getId()).createFile(blob);

    var url = driveFile.getDownloadUrl();

    if(url.match(/\?e=download&gd=true/i)) {
      url = driveFile.getDownloadUrl().replace(/\?e=download&gd=true/i,''); //getDownloadUrl()で得られるurlの末尾に、デフォルトで不用なパラメータが付いてしまうので、削減している。
    } else if(url.match(/&e=download&gd=true/i)){
      url = driveFile.getDownloadUrl().replace(/&e=download&gd=true/i,'');
    }
    return url;
  },

  // スプレッドシート名では重複しないように末尾に(en)/(ja)をつけて分類しているが、実際に使用するファイルには必要ないため、ダウンロードできる形にする前に削減している。
  getFileNameForDownload: function(name) {
    if(name.match(/en/i)) {
      return name.replace(/\(en\)/i, "");

    } else if(name.match(/ja/i)) {
      return name.replace(/\(ja\)/i, "");

    }
  },

  displayDownloadModalView: function(title) {
    const html = HtmlService.createTemplateFromFile(RESOURCES.HTML.MODAL).evaluate().setHeight(300).setWidth(300);
    SpreadsheetApp.getUi().showModalDialog(html, title);
  }

}


/* ----------------- Callback　Functions ----------------- */
LOCALIZED_STRINGS.CONTROLLERS.CALLBACK = {

  getReplacementStringForIos: function(index, message) {
    const isMatch = message.match(/\(\(.*?\)\)/);
    if(!isMatch) { return '%'+index+'$@'; }

    const msgString = isMatch.toString();
    const onlyMessage = this.getSubstr(msgString);
    return '%' + onlyMessage + '$@';
  },

  getReplacementStringForAndroid: function(index, message) {
    var id = index;
    var type = RESOURCES.TYPE_ANDROID.STRING;

    var isMatch = message.match(/\{\{.*?\}\}/);
    if(isMatch) {
      const onlyMessage = this.getSubstr(isMatch.toString());
      if(Number(onlyMessage)) {
        type = RESOURCES.TYPE_ANDROID.INT;
      }
    }

    var isMatch = message.match(/\(\(.*?\)\)/);
    if(isMatch) {
      id = this.getSubstr(isMatch.toString());
    }

    return '%' + id + '$' + type;
  },

  getSubstr: function(str) {
    return str.substr(2).substr(0, str.length-4); //　最初と最後の括弧、4文字分を引いている。
  },

  getLineCharacterForIos: function(id, message) {
    return ['"'+ id + '"' + '="' + message + '";'];
  },

  getLineCharacterForAndroid: function(id, message) {
    return ['    <string name="' + id + '">' + message + '</string>']; //インデントのために意図的に前に半角スペースが4つ入っている。;
  },
}


/* ----------------- Custom Menu　----------------- */
var CUSTOM_MENU = CUSTOM_MENU || {};

CUSTOM_MENU = {
  MENU_TITLE: "Custom Menu",
  MENU_CONTENTS: [
    { name: "Convert MessageList to Localized File", functionName: "generateMessageListStringsFiles" },
    { name: "Remove All Localized Files", functionName: "removeStringsFiles" },
    { name: "Download Localized Files For Ios", functionName: "downloadStringsFilesForIos" },
    { name: "Download Localized Files For Android", functionName: "downloadStringsFilesForAndroid" }
  ],

  add: function() {
    SpreadsheetApp.getActiveSpreadsheet().addMenu(this.MENU_TITLE, this.MENU_CONTENTS);
  }
}



/* --------------------------------- Functions --------------------------------- */

/* ----------------- Menu Functions ---------------- */
/* Custom Menu Function */
function onOpen() { CUSTOM_MENU.add(); };


/* ----------------- Localized Functions ---------------- */
/* SetUp */
LOCALIZED_STRINGS.init();

/* Generate Localized File Function */
function generateMessageListStringsFiles() { LOCALIZED_STRINGS.generate(); };

/* Remove Function */
function removeStringsFiles() { LOCALIZED_STRINGS.remove(); };

/* Download Function */
function downloadStringsFilesForIos() { LOCALIZED_STRINGS.downloadForIos(); };
function downloadStringsFilesForAndroid() { LOCALIZED_STRINGS.downloadForAndroid(); };


/* ----------------- Common Functions ---------------- */
/* Log Function */
function log(str) { Logger.log(str); };
