(function(jQuery, XLSX) {

  'use strict';

  var Extractor = function() {

    this.reader = null;
    this.useWorker = true;

    this.fileElement = '#xlf';
    this.dropElement = '#drop';
    this.useWorkerElement = '#useWorker';
    this.filterDuplicates = '#filterDownloads';

    this.uploadContainer = '.upload-container';
    this.loadingContainer = '.loading-container';
    this.resultContainer = '.result-container';
    this.resultHeader = '.result-header';

    this.workerScript = './js/xlsxworker.js';
  };

  Extractor.prototype.init = function() {
    this.reader = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
    this.useWorker = jQuery(this.useWorkerElement).is(':checked');
    this.bindEvents();
  };

  Extractor.prototype.bindEvents = function() {

    jQuery(this.fileElement).on('change', this.onFileChange.bind(this));
    jQuery(this.dropElement).on('drop', this.onFileDrop.bind(this));
    jQuery(this.dropElement).on('dragenter', this.onDragOver.bind(this));
    jQuery(this.dropElement).on('dragover', this.onDragOver.bind(this));
  };

  Extractor.prototype.readFile = function(file) {

    var reader = new FileReader();
    reader.addEventListener("load", this.onLoadFileReader.bind(this), false);

    if (this.reader) {
      reader.readAsBinaryString(file)
    } else {
      reader.readAsArrayBuffer(file)
    }
  };

  Extractor.prototype.onLoadFileReader = function(event) {

    if (this.useWorker) {
      return this.processWorkers(event.target.result, this.processWorkbook);
    }

    var workbook = null;
    if (this.reader) {
      workbook = XLSX.read(event.target.result, {
        type: 'binary'
      });
    } else {
      workbook = XLSX.read(btoa(this.fixData(event.target.result)), {
        type: 'base64'
      });
    }

    return this.processWorkbook(workbook);
  };

  Extractor.prototype.onDragOver = function(event) {

    event.stopPropagation();
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  };

  Extractor.prototype.onFileChange = function(event) {
    var files = event.target.files;
    if (files.length) {
      this.readFile(files[0]);
    }
  };

  Extractor.prototype.onFileDrop = function(event) {
    console.log(event);
  };

  Extractor.prototype.iterateSheets = function(workbook, callback) {

    workbook.SheetNames.forEach(function(sheetName) {
      callback.call(this, workbook.Sheets[sheetName], sheetName);
    }.bind(this));
  };

  Extractor.prototype.processWorkbook = function(workbook) {

    this.iterateSheets(workbook, this.fileToCSV.bind(this));
  };

  Extractor.prototype.fileToCSV = function(sheet, sheetName) {

    var csv = XLSX.utils.sheet_to_csv(sheet);
    if (csv.length) {
      this.formSheetData(csv, sheetName);
    }
  };

  Extractor.prototype.formSheetData = function(csv, sheetName) {

    var em = this.extractEmailsFromString(csv);
    var listTemplate = this.createListTemplate(em);
    var panelTemplate = this.sheetPanelTemplate(sheetName, Object.keys(em).length, listTemplate);
    jQuery(this.resultContainer).append(panelTemplate);

    this.showResults();
  };

  Extractor.prototype.processWorkers = function(data, callback) {

    var worker = new Worker(this.workerScript);
    worker.addEventListener("message", this.onWorkerMessage.bind(this, callback));

    var arr = this.reader ? data : btoa(this.fixData(event.target.result));
    worker.postMessage({
      d: arr,
      b: this.reader
    });
  };

  Extractor.prototype.onWorkerMessage = function(callback, event) {

    switch (event.data.t) {
      case "ready":
        break;
      case "start":
        break;
      case "end":
        break;
      case "e":
        console.error(event.data.d);
        break;
      case "xlsx":
        callback.call(this, JSON.parse(event.data.d));
      default:
        break;
    };
  };

  Extractor.prototype.extractEmailsFromString = function(text) {

    var matches = text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
    if (!matches)
      return [];

    return this.removeDuplicateValues(matches);
  };

  Extractor.prototype.removeDuplicateValues = function(values) {

    return values.filter(function(item, index, inputArray) {
      return inputArray.indexOf(item) == index;
    });
  };

  Extractor.prototype.fixData = function(data) {

    var o = "",
      l = 0,
      w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
  };

  Extractor.prototype.downloadToCSV = function(values) {

    if(jQuery(this.filterDuplicates).is(':checked')) {
      values = this.removeDuplicateValues(values);
    }

    var hiddenElement = document.createElement('a');
    hiddenElement.href = 'data:text/csv;charset=utf-8,' + encodeURI(values);
    hiddenElement.target = '_blank';
    hiddenElement.download = Date.now() + '.csv';
    hiddenElement.click();
  };

  Extractor.prototype.sheetPanelTemplate = function(name, count, data) {

    return '<div class="col-md-4">' +
      '<div class="panel panel-default">' +
      '<div class="panel-heading">' + name +
      '<span class="badge pull-right">' + count + '</span>' +
      '</div>' +
      '<div class="panel-body no-padding">' +
      '<ul class="list-group no-side-borders no-margin">' +
      '' + data.join('') +
      '</ul>' +
      '</div>' +
      '<div class="panel-footer">' +
      '<a href="#" class="btn btn-default">' +
      'Download Sheet' +
      '</a>' +
      '</div>' +
      '</div>' +
      '</div>  ';
  };

  Extractor.prototype.createListTemplate = function(items) {
    return jQuery.map(items, this.listItemTemplate.bind(this));
  };

  Extractor.prototype.listItemTemplate = function(item) {
    return '<li class="list-group-item">' + item + '</li>';
  };

  Extractor.prototype.showResults = function() {
    jQuery(this.resultHeader).show();
    jQuery(this.resultContainer).show();
  };

  Extractor.prototype.HideResults = function() {
    jQuery(this.resultHeader).hide();
    jQuery(this.resultContainer).hide();
  };

  var app = new Extractor();
  app.init();

})(jQuery, XLSX)
