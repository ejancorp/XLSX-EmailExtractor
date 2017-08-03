    /*jshint browser:true */
    /* eslint-env browser */
    /* eslint no-use-before-define:0 */
    /*global Uint8Array, Uint16Array, ArrayBuffer */
    /*global XLSX */
    var X = XLSX;
    var XW = {
      /* worker message */
      msg: 'xlsx',
    };

    var rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";

    var wtf_mode = false;

    function fixdata(data) {
      var o = "",
        l = 0,
        w = 10240;
      for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
      return o;
    }

    function ab2str(data) {
      var o = "",
        l = 0,
        w = 10240;
      for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l * w, l * w + w)));
      o += String.fromCharCode.apply(null, new Uint16Array(data.slice(l * w)));
      return o;
    }

    function s2ab(s) {
      var b = new ArrayBuffer(s.length * 2),
        v = new Uint16Array(b);
      for (var i = 0; i != s.length; ++i) v[i] = s.charCodeAt(i);
      return [v, b];
    }

    function get_radio_value(radioName) {
      var radios = document.getElementsByName(radioName);
      for (var i = 0; i < radios.length; i++) {
        if (radios[i].checked || radios.length === 1) {
          return radios[i].value;
        }
      }
    }

    function to_json(workbook) {
      var result = {};
      workbook.SheetNames.forEach(function(sheetName) {
        var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName]);
        if (roa.length > 0) {
          result[sheetName] = roa;
        }
      });
      return result;
    }

    function to_csv(workbook) {
      var result = [];
      workbook.SheetNames.forEach(function(sheetName) {
        var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
        if (csv.length > 0) {
          result.push("SHEET: " + sheetName);
          result.push("");
          result.push(csv);
        }
      });
      return result.join("\n");
    }

    function to_formulae(workbook) {
      var result = [];
      workbook.SheetNames.forEach(function(sheetName) {
        var formulae = X.utils.get_formulae(workbook.Sheets[sheetName]);
        if (formulae.length > 0) {
          result.push("SHEET: " + sheetName);
          result.push("");
          result.push(formulae.join("\n"));
        }
      });
      return result.join("\n");
    }

    var HTMLOUT = document.getElementById('htmlout');
    var OUT = document.getElementById('out');

    function to_html(workbook) {
      HTMLOUT.innerHTML = "";
      OUT.innerHTML = "";
      workbook.SheetNames.forEach(function(sheetName) {
        var htmlstr = X.write(workbook, {
          sheet: sheetName,
          type: 'binary',
          bookType: 'html'
        });
        HTMLOUT.innerHTML += htmlstr;
      });
    }

    var tarea = document.getElementById('b64data');

    function b64it() {
      if (typeof console !== 'undefined') console.log("onload", new Date());
      var wb = X.read(tarea.value, {
        type: 'base64',
        WTF: wtf_mode
      });
      process_wb(wb);
    }
    window.b64it = b64it;

    var global_wb;

    function process_wb(wb) {
      global_wb = wb;
      var output = "";
      switch (get_radio_value("format")) {
        case "json":
          output = JSON.stringify(to_json(wb), 2, 2);
          break;
        case "form":
          output = to_formulae(wb);
          break;
        case "html":
          return to_html(wb);
        default:
          output = to_csv(wb);
      }
      if (OUT.innerText === undefined) OUT.textContent = output;
      else OUT.innerText = output;
      if (typeof console !== 'undefined') console.log("output", new Date());
    }

    function setfmt() {
      if (global_wb) process_wb(global_wb);
    }
    window.setfmt = setfmt;

    var drop = document.getElementById('drop');

    function handleDrop(e) {
      e.stopPropagation();
      e.preventDefault();
      var files = e.dataTransfer.files;
      var f = files[0]; {
        var reader = new FileReader();
        //var name = f.name;
        reader.onload = function(e) {
          if (typeof console !== 'undefined') console.log("onload", new Date());
          var data = e.target.result;

          var wb;
          if (rABS) {
            wb = X.read(data, {
              type: 'binary'
            });
          } else {
            var arr = fixdata(data);
            wb = X.read(btoa(arr), {
              type: 'base64'
            });
          }
          process_wb(wb);

        };
        if (rABS) reader.readAsBinaryString(f);
        else reader.readAsArrayBuffer(f);
      }
    }

    function handleDragover(e) {
      e.stopPropagation();
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    }

    if (drop.addEventListener) {
      drop.addEventListener('dragenter', handleDragover, false);
      drop.addEventListener('dragover', handleDragover, false);
      drop.addEventListener('drop', handleDrop, false);
    }


    var xlf = document.getElementById('xlf');

    function handleFile(e) {
      var files = e.target.files;
      var f = files[0]; {
        var reader = new FileReader();
        //var name = f.name;
        reader.onload = function(e) {
          if (typeof console !== 'undefined') console.log("onload", new Date(), rABS);
          var data = e.target.result;
          var wb;
          if (rABS) {
            wb = X.read(data, {
              type: 'binary'
            });
          } else {
            var arr = fixdata(data);
            wb = X.read(btoa(arr), {
              type: 'base64'
            });
          }
          process_wb(wb);

        };
        if (rABS) reader.readAsBinaryString(f);
        else reader.readAsArrayBuffer(f);
      }
    }

    var fs, err = function(e) {
      throw e;
    };;
    window.requestFileSystem = window.requestFileSystem || window.webkitRequestFileSystem;
    window.requestFileSystem(window.TEMPORARY, 5 * 1024 * 1024, initFS, errorHandler);

    var filesInput = document.getElementById("files");

    function initFS(_fs) {
      _fs.root.getDirectory('Documents', {
        create: true
      }, function(dirEntry) {
        console.log(dirEntry)
        fs = _fs;
      }, errorHandler);
    }

    function errorHandler(msg) {
      console.log('An error occured: ' + msg);
    }

    filesInput.addEventListener('change', onDirectorySelect, false)

    function onDirectorySelect(event) {
      var files = this.files;
      if (!files) return;

      $.each(files, function(key, value) {
        saveFile(value);
      })
    }

    function saveFile(file) {

      var text = file ? file.name : "file-" + Date.now();

      if (!file) return;

      // create a sandboxed file
      fs.root.getFile(
        'Documents/' + file.name, {
          create: true
        },
        function(fileEntry) {
          // create a writer that can put data in the file
          fileEntry.createWriter(function(writer) {
            writer.onwriteend = function(fileDone) {
              getFilesDirectory()
            };
            writer.onerror = err;

            // this will read the contents of the current file
            var fr = new FileReader;
            fr.onloadend = function() {
              // create a blob as that's what the
              // file writer wants
              var builder = new Blob([fr.result]);
              writer.write(builder);
            };
            fr.onerror = err;
            fr.readAsArrayBuffer(file);
          }, err);
        },
        err
      );
    }

    var fileList = [];

    function getFilesDirectory() {
      fs.root.getDirectory('Documents', {
        create: false
      }, function(dirEntry) {
        var reader = dirEntry.createReader();
        reader.readEntries(function(results) {
          fileList = results
          displayFiles(results);
        });
      }, errorHandler);
    }

    function displayFiles(results) {
      $.each(results, function(key, file) {

        fs.root.getFile(
          'Documents/' + file.name, {
            create: true
          },
          function(fileEntry) {
            fileEntry.getMetadata(function(metadata) {
              console.log(metadata)
              $('.list-group').append('<a href="#" class="list-group-item">' + file.name + '<span class="badge">' + bytesToSize(metadata.size) + '</span></a>')
            })
          },
          err
        );

      });
    }


    function bytesToSize(bytes) {
      var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
      if (bytes == 0) return 'n/a';
      var i = parseInt(Math.floor(Math.log(bytes) / Math.log(1024)));
      if (i == 0) return bytes + ' ' + sizes[i];
      return (bytes / Math.pow(1024, i)).toFixed(1) + ' ' + sizes[i];
    };

    if (xlf.addEventListener) xlf.addEventListener('change', handleFile, false);
