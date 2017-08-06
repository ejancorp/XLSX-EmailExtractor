(function(jQuery, XLSX, Modal) {

  'use strict';

  var Extractor = function() {

    this.file = null;

    this.reader = null;
    this.useWorker = true;

    this.fileElement = '#xlf';
    this.dropElement = '#drop';
    this.useWorkerElement = '#useWorker';
    this.filterDuplicates = '#filterDownloads';

    this.downloadAllButton = '#downloadAll';
    this.downloadSheetButton = '.downloadSheet';
    this.allCount = '#allCount';

    this.uploadContainer = '.upload-container';
    this.loadingContainer = '.loading-container';
    this.resultContainer = '.result-container';
    this.resultHeader = '.result-header';
    this.panelTitle = '.sheet-title';
    this.sheetCount = '.sheet-count';

    this.itemClass = 'extracted-item';
    this.panelClass = 'extracted-panel';

    this.workerScript = './js/xlsxworker.js';

    this.allowedFileType = ['xlsx', 'csv'];

    this.els = [];
    this.sheetCount = 0;

    this.collection = [];
  };

  Extractor.prototype.init = function() {

    try {
      var worker = new Worker(this.workerScript);
    } catch (e) {

      jQuery(this.useWorkerElement).attr('disabled', true);
      jQuery(this.useWorkerElement).prop('checked', false);
      jQuery(this.useWorkerElement).parents('.checkbox').hide();
      this.useWorker = false;

      console.log(e);
    }

    this.reader = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";
    this.bindEvents();
  };

  Extractor.prototype.bindEvents = function() {

    jQuery(this.fileElement).on('change', this.onFileChange.bind(this));
    jQuery(this.dropElement).on('drop', this.onFileDrop.bind(this));
    jQuery(this.dropElement).on('dragenter', this.onDragOver.bind(this));
    jQuery(this.dropElement).on('dragover', this.onDragOver.bind(this));

    jQuery(this.downloadAllButton).on('click', this.downloadAll.bind(this));
    jQuery(document).on('click', this.downloadSheetButton, this.downloadSheet.bind(this));
    jQuery(this.filterDuplicates).on('click', this.renderAllCount.bind(this));
  };

  Extractor.prototype.readFile = function(file) {

    if (jQuery.inArray(String(file.name).split('.').pop(), this.allowedFileType) == -1) {
      Modal.alert('File type is not allowed. Allowed file type/s: ' + this.allowedFileType.join(', '));
      return false;
    }

    this.file = file;

    var reader = new FileReader();
    reader.addEventListener("load", this.onLoadFileReader.bind(this), false);
    reader.addEventListener("progress", this.onProgressFileReader.bind(this), false);

    if (this.reader) {
      reader.readAsBinaryString(file)
    } else {
      reader.readAsArrayBuffer(file)
    }
  };

  Extractor.prototype.onProgressFileReader = function(event) {
    console.log(event);
  };

  Extractor.prototype.onLoadFileReader = function(event) {

    this.startProgress();
    this.HideResults();
    jQuery(this.resultContainer).empty();

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

    event = event.originalEvent ? event.originalEvent : event;
    event.dataTransfer.dropEffect = 'copy';
  };

  Extractor.prototype.onFileChange = function(event) {

    this.useWorker = jQuery(this.useWorkerElement).is(':checked');

    var files = event.target.files;
    if (files.length) {
      this.readFile(files[0]);
    }
  };

  Extractor.prototype.onFileDrop = function(event) {

    event.stopPropagation();
    event.preventDefault();

    event = event.originalEvent ? event.originalEvent : event;

    this.useWorker = jQuery(this.useWorkerElement).is(':checked');

    var files = event.dataTransfer.files;
    if (files.length) {
      this.readFile(files[0]);
    }
  };

  Extractor.prototype.iterateSheets = function(workbook, callback) {
    this.sheetCount = Object.keys(workbook.SheetNames).length;
    workbook.SheetNames.forEach(function(sheetName) {
      callback.call(this, workbook.Sheets[sheetName], sheetName);
    }.bind(this));
  };

  Extractor.prototype.processWorkbook = function(workbook) {

    jQuery(this.resultContainer).empty();

    this.iterateSheets(workbook, this.fileToCSV.bind(this));
  };

  Extractor.prototype.fileToCSV = function(sheet, sheetName) {

    var csv = JSON.stringify(sheet);
    if (csv) {
      this.formSheetData(csv, sheetName);
    }
  };

  Extractor.prototype.formSheetData = function(csv, sheetName) {

    var em = this.extractEmailsFromString(csv);
    var listTemplate = this.createListTemplate(em);
    var panelTemplate = this.sheetPanelTemplate(sheetName, Object.keys(em).length, listTemplate);

    this.collection.push(em);
    this.els.push(panelTemplate);
    this.sheetCount -= 1;

    if(this.sheetCount <= 0) {
        jQuery(this.resultContainer).html(this.els.join(''));
        this.showResults();
        this.renderAllCount();
        this.stopProgress();
    }
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

  Extractor.prototype.sheetPanelTemplate = function(name, count, data) {

    return '<div class="col-md-4">' +
      '<div class="panel panel-default ' + this.panelClass + '">' +
      '<div class="panel-heading"><b class="sheet-title">' + name + '</b>' +
      '<span class="badge sheet-count pull-right">' + count + '</span>' +
      '</div>' +
      '<div class="panel-body no-padding">' +
      '<ul class="list-group no-side-borders no-margin">' +
      '' + data.join('') +
      '</ul>' +
      '</div>' +
      '<div class="panel-footer">' +
      '<a href="#" class="btn btn-default downloadSheet">' +
      'Download Sheet' +
      '</a>' +
      '</div>' +
      '</div>' +
      '</div>';
  };

  Extractor.prototype.renderAllCount = function() {

    var combine = [];
    jQuery.each(this.collection, function(key, value) {
      jQuery.each(value, function(k, v) {
        combine.push(v);
      });
    });

    if (jQuery(this.filterDuplicates).is(':checked')) {
      combine = this.removeDuplicateValues(combine);
    }

    jQuery(this.allCount).text(combine.length);
  };

  Extractor.prototype.createListTemplate = function(items) {
    return jQuery.map(items, this.listItemTemplate.bind(this));
  };

  Extractor.prototype.listItemTemplate = function(item) {
    return '<li class="list-group-item ' + this.itemClass + '">' + item + '</li>';
  };

  /**
   * Show results container
   * @return {Void}
   */
  Extractor.prototype.showResults = function() {
    jQuery(this.resultHeader).show();
    jQuery(this.resultContainer).show();
  };

  /**
   * Hide results container
   * @return {Void}
   */
  Extractor.prototype.HideResults = function() {
    jQuery(this.resultHeader).hide();
    jQuery(this.resultContainer).hide();
  };

  /**
   * Show progress container
   * @return {Void}
   */
  Extractor.prototype.startProgress = function() {
    jQuery(this.uploadContainer).hide();
    jQuery(this.loadingContainer).show();
  };

  /**
   * Hide progress container
   * @return {Void}
   */
  Extractor.prototype.stopProgress = function() {
    jQuery(this.uploadContainer).show();
    jQuery(this.loadingContainer).hide();
  };

  /**
   * Event function on click button, get related items and create an array
   * @param  {Object} event Event object
   * @return {Void}
   */
  Extractor.prototype.downloadAll = function(event) {

    var items = jQuery.map(jQuery('.' + this.itemClass), this.getItemString.bind(this));
    this.downloadToCSV(this.file.name, items);
    event.preventDefault();
  };

  /**
   * Event function on click button, get related items and create an array
   * @param  {Object} event Event object
   * @return {Void}
   */
  Extractor.prototype.downloadSheet = function(event) {

    var panel = jQuery(event.target).parents('.' + this.panelClass);
    var items = jQuery.map(panel.find('.' + this.itemClass), this.getItemString.bind(this));

    this.downloadToCSV([panel.find(this.panelTitle).text(), this.file.name].join('-'), items);
    event.preventDefault();
  };

  /**
   * Get element text or string
   * @param  {String} value Element string
   * @return {String}       string without html tags
   */
  Extractor.prototype.getItemString = function(value) {
    return jQuery(value).text();
  };

  /**
   * download array as csv file
   * @param  {String} name  prefix name for csv
   * @param  {Array} values list of data in array
   * @return {Void}
   */
  Extractor.prototype.downloadToCSV = function(name, values) {

    name = name ? name.replace(/ /g,"-") : 'file';

    if (jQuery(this.filterDuplicates).is(':checked')) {
      values = this.removeDuplicateValues(values);
    }

    if (jQuery.isArray(values)) {
      values = values.join('\n');
    }

    var hiddenElement = document.createElement('a');
    hiddenElement.href = 'data:text/csv;charset=utf-8,' + encodeURI(values);
    hiddenElement.target = '_blank';
    hiddenElement.download = name + '-' + Date.now() + '.csv';
    hiddenElement.click();
  };

  var app = new Extractor();
  app.init();

})(jQuery, XLSX, bootbox);
