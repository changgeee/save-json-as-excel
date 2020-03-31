/*
 * @Author: changge
 * @Date: 2017-08-15 11:07:05
 * @Last Modified by: changge
 * @Last Modified time: 2018-01-22 17:49:31
 */
(function(global) {
  'use strict';
  var Base64 = {
    // private property
    _keyStr: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=',
    // public method for encoding
    encode: function(input) {
      var output = '';
      var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
      var i = 0;

      input = Base64._utf8_encode(input);

      while (i < input.length) {
        chr1 = input.charCodeAt(i++);
        chr2 = input.charCodeAt(i++);
        chr3 = input.charCodeAt(i++);

        enc1 = chr1 >> 2;
        enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
        enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
        enc4 = chr3 & 63;

        if (isNaN(chr2)) {
          enc3 = enc4 = 64;
        } else if (isNaN(chr3)) {
          enc4 = 64;
        }

        output =
          output +
          this._keyStr.charAt(enc1) +
          this._keyStr.charAt(enc2) +
          this._keyStr.charAt(enc3) +
          this._keyStr.charAt(enc4);
      } // Whend

      return output;
    }, // End Function encode

    // public method for decoding
    decode: function(input) {
      var output = '';
      var chr1, chr2, chr3;
      var enc1, enc2, enc3, enc4;
      var i = 0;

      input = input.replace(/[^A-Za-z0-9\+\/\=]/g, '');
      while (i < input.length) {
        enc1 = this._keyStr.indexOf(input.charAt(i++));
        enc2 = this._keyStr.indexOf(input.charAt(i++));
        enc3 = this._keyStr.indexOf(input.charAt(i++));
        enc4 = this._keyStr.indexOf(input.charAt(i++));

        chr1 = (enc1 << 2) | (enc2 >> 4);
        chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
        chr3 = ((enc3 & 3) << 6) | enc4;

        output = output + String.fromCharCode(chr1);

        if (enc3 != 64) {
          output = output + String.fromCharCode(chr2);
        }

        if (enc4 != 64) {
          output = output + String.fromCharCode(chr3);
        }
      } // Whend

      output = Base64._utf8_decode(output);

      return output;
    }, // End Function decode

    // private method for UTF-8 encoding
    _utf8_encode: function(string) {
      var utftext = '';
      string = string.replace(/\r\n/g, '\n');

      for (var n = 0; n < string.length; n++) {
        var c = string.charCodeAt(n);

        if (c < 128) {
          utftext += String.fromCharCode(c);
        } else if (c > 127 && c < 2048) {
          utftext += String.fromCharCode((c >> 6) | 192);
          utftext += String.fromCharCode((c & 63) | 128);
        } else {
          utftext += String.fromCharCode((c >> 12) | 224);
          utftext += String.fromCharCode(((c >> 6) & 63) | 128);
          utftext += String.fromCharCode((c & 63) | 128);
        }
      } // Next n

      return utftext;
    }, // End Function _utf8_encode

    // private method for UTF-8 decoding
    _utf8_decode: function(utftext) {
      var string = '';
      var i = 0;
      var c, c1, c2, c3;
      c = c1 = c2 = 0;

      while (i < utftext.length) {
        c = utftext.charCodeAt(i);

        if (c < 128) {
          string += String.fromCharCode(c);
          i++;
        } else if (c > 191 && c < 224) {
          c2 = utftext.charCodeAt(i + 1);
          string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
          i += 2;
        } else {
          c2 = utftext.charCodeAt(i + 1);
          c3 = utftext.charCodeAt(i + 2);
          string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
          i += 3;
        }
      } // Whend

      return string;
    }, // End Function _utf8_decode
  };
  var saveJsonAsExcel = (function() {
    var jsonToXLS = (function() {
      var xlsTemp =
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=ProgId content=Excel.Sheet> <meta name=Generator content="Microsoft Excel 11"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>${table}</table></body></html>';
      return function(dataSource, columns) {
        var xlsData = '',
          values = [];
        xlsData += '<thead>';
        columns.forEach(function(item) {
          xlsData += '<th>' + item.title + '</th>';
          if (item.computed && typeof item.computed === 'function') {
            values.push(item.computed);
          } else if (item.dataIndex && typeof item.dataIndex === 'string') {
            values.push(item.dataIndex);
          } else {
            values.push(null);
          }
        });
        xlsData += '</thead>';
        xlsData += '<tbody>';
        dataSource.forEach(function(item) {
          xlsData += '<tr>';
          values.forEach(function(key) {
            xlsData += '<td>';
            if (typeof key === 'string') {
              xlsData += item[key];
            } else if (typeof key === 'function') {
              xlsData += key(item) || '';
            } else {
              xlsData += '';
            }
            xlsData += '</td>';
          });
          xlsData += '</tr>';
        });
        xlsData += '</tbody>';
        return xlsTemp.replace('${table}', xlsData);
      };
    })();

    var jsonToCSV = function(dataSource, columns) {
      var csvData = '',
        values = [];
      columns.forEach(function(item) {
        csvData += item.title;
        csvData += ',';
        if (item.computed && typeof item.computed === 'function') {
          values.push(item.computed);
        } else if (item.dataIndex && typeof item.dataIndex === 'string') {
          values.push(item.dataIndex);
        } else {
          values.push(null);
        }
      });
      csvData = csvData.slice(0, csvData.length - 1);
      csvData += '\r\n';
      dataSource.forEach(function(item) {
        values.forEach(function(key) {
          if (typeof key === 'string') {
            csvData += item[key];
          } else if (typeof key === 'function') {
            csvData += key(item) || '';
          } else {
            csvData += '';
          }
          csvData += ',';
        });
        csvData = csvData.slice(0, csvData.length - 1);
        csvData += '\r\n';
      });
      return csvData;
    };
    var base64 = function(s) {
      return Base64.encode('\ufeff' + s);
    };
    var saveAsXls = function(dataSource, columns, fileName) {
      download(jsonToCSV(dataSource, columns), 'application/vnd.ms-excel', fileName);
    };
    var saveAsCsv = function(dataSource, columns, fileName) {
      download(jsonToCSV(dataSource, columns), 'text/csv', fileName);
    };

    var download = function(data, mime, fileName) {
      if (window.navigator.msSaveBlob) {
        var blob = dataToBlob(data, mime);
        window.navigator.msSaveBlob(blob, fileName);
        return false;
      }
      var alink = document.createElement('a');
      if (window.URL.createObjectURL) {
        var blob = dataToBlob(data, mime);
        var blobUrl = window.URL.createObjectURL(blob);
        alink.href = blobUrl;
        alink.download = fileName;
        alink.click();
        return;
      }
      if (alink.download === '') {
        alink.href = 'data:' + mime + ';base64,' + base64(data);
        alink.download = fileName;
        alink.click();
        return;
      }
    };
    var dataToBlob = function(data, mime) {
      var bstr = '\ufeff' + data;
      return new Blob([bstr], {
        type: mime,
      });
    };
    return {
      saveAsXls: saveAsXls,
      saveAsCsv: saveAsCsv,
    };
  })();
  if (typeof define === 'function' && define.amd) {
    define(function() {
      return saveJsonAsExcel;
    });
  } else if (typeof exports !== 'undefined') {
    if (typeof module !== 'undefined' && module.exports) {
      exports = module.exports = saveJsonAsExcel;
    }
    exports.saveJsonAsExcel = saveJsonAsExcel;
  } else {
    global.saveJsonAsExcel = saveJsonAsExcel;
  }
})(this);
