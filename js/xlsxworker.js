importScripts('../lib/cpexcel.js');
importScripts('../lib/jszip.js');
importScripts('../lib/xlsx.js');
postMessage({
  t: "ready"
});

onmessage = function(oEvent) {
  var v;
  try {
    v = XLSX.read(oEvent.data.d, {
      type: oEvent.data.b ? 'binary' : 'base64'
    });
    postMessage({
      t: "start"
    });
  } catch (e) {
    postMessage({
      t: "e",
      d: e.stack || e
    });
    postMessage({
      t: "end"
    });
  }
  postMessage({
    t: "xlsx",
    d: JSON.stringify(v)
  });
  postMessage({
    t: "end"
  });
};
