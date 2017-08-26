/*
 * @Author: changge
 * @Date: 2017-08-15 11:07:05
 * @Last Modified by: changge
 * @Last Modified time: 2017-08-26 20:49:40
 */
(function (global) {
  'use strict';
  var JSONTOEXCEL = (function () {
    var jsonToXLS = (function () {
      var xlsTemp = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=ProgId content=Excel.Sheet> <meta name=Generator content="Microsoft Excel 11"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>${table}</table></body></html>'
      return function (data, header) {
        var xlsData = '', keys = []
        if (header) {
          xlsData += '<thead>'
          for (var key in header) {
            keys.push(key)
            xlsData += '<th>' + header[key] + '</th>'
          }
          xlsData += '</thead>'
          xlsData += '<tbody>'
          data.map(function (item, index) {
            xlsData += '<tr>'
            for (var i = 0; i < keys.length; i++) {
              xlsData += '<td>' + item[keys[i]] + '</td>'
            }
            xlsData += '</tr>'
          })
          xlsData += '</tbody>'
          return xlsTemp.replace('${table}', xlsData)
        }
        data.map(function (item, index) {
          xlsData += '<tbody><tr>'
          for (var key in item) {
            xlsData += '<td>' + item[key] + '</td>'
          }
          xlsData += '</tr></tbody>'
        })
        return xlsTemp.replace('${table}', xlsData)
      }
    })()


    var jsonToCSV = function (data, keys) {
      var csvData = ''
      if (keys) {
        data.map(function (item) {
          for (var i = 0; i < keys.length; i++) {
            csvData += item[keys[i]] + ','
          }
          csvData = csvData.slice(0, csvData.length - 1)
          csvData += '\r\n'
        })
        return csvData
      }
      data.map(function (item) {
        for (var k in item) {
          csvData += item[k] + ','
        }
        csvData = csvData.slice(0, csvData.length - 1)
        csvData += '\r\n'
      })
      return csvData
    }
    var base64 = function (s) {
      return window.btoa(window.unescape(encodeURIComponent(s)))
    }
    var exportXLS = function (data, fileName, header) {
      var XLSData = 'data:application/vnd.ms-excel;base64,' + base64(jsonToXLS(data, header))
      download(XLSData, fileName)
    }
    var exportCSV = function (data, fileName, keys) {
      var CSVData = 'data:application/csv;base64,' + base64(jsonToCSV(data, keys))
      download(CSVData, fileName)
    }

    var download = function (base64data, fileName) {
      if (window.navigator.msSaveBlob) {
        var blob = base64ToBlob(base64data)
        window.navigator.msSaveBlob(blob, filename)
        return false;
      }
      var alink = document.createElement('a')
      if (window.URL.createObjectURL) {
        var blob = base64ToBlob(base64data)
        var blobUrl = window.URL.createObjectURL(blob)
        alink.href = blobUrl
        alink.download = fileName
        alink.click()
        return
      }
      if (alink.download === '') {
        alink.href = base64data
        alink.download = fileName
        alink.click()
        return
      }
    }
    var base64ToBlob = function (base64Data) {
      var arr = base64Data.split(','), mime = arr[0].match(/:(.*?);/)[1],
        bstr = atob(arr[1]), n = bstr.length, u8arr = new Uint8Array(n)
      while (n--) {
        u8arr[n] = bstr.charCodeAt(n)
      }
      return new Blob([u8arr], { type: mime })
    }
    return {
      exportXLS: exportXLS,
      exportCSV: exportCSV
    }
  }())
  if (typeof define === 'function' && define.amd) {
    define(function () { return JSONTOEXCEL })
  } else if (typeof exports !== 'undefined') {
    if (typeof module !== 'undefined' && module.exports) {
      exports = module.exports = JSONTOEXCEL
    }
    exports.JSONTOEXCEL = JSONTOEXCEL
  } else {
    global.JSONTOEXCEL = JSONTOEXCEL
  }
})(this)