/*
 * @Author: changge
 * @Date: 2017-08-15 11:07:05
 * @Last Modified by: changge
 * @Last Modified time: 2018-01-22 17:49:31
 */
(function (global) {
  'use strict';
  var JSON2EXCEL = (function () {
    var jsonToXLS = (function () {
      var xlsTemp = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta name=ProgId content=Excel.Sheet> <meta name=Generator content="Microsoft Excel 11"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>${table}</table></body></html>'
      return function (dataSource, columns) {
        var xlsData = '',
          values = []
        xlsData += '<thead>'
        columns.forEach(function (item) {
          xlsData += '<th>' + item.title + '</th>'
          if (item.computed && typeof item.computed === 'function') {
            values.push(item.computed)
          } else if (item.dataIndex && typeof item.dataIndex === 'string') {
            values.push(item.dataIndex)
          } else {
            values.push(null)
          }
        })
        xlsData += '</thead>'
        xlsData += '<tbody>'
        dataSource.forEach(function (item) {
          xlsData += '<tr>'
          values.forEach(function (key) {
            xlsData += '<td>'
            if (typeof key === 'string') {
              xlsData += item[key]
            } else if (typeof key === 'function') {
              xlsData += key(item) || ''
            } else {
              xlsData += ''
            }
            xlsData += '</td>'
          })
          xlsData += '</tr>'
        })
        xlsData += '</tbody>'
        return xlsTemp.replace('${table}', xlsData)
      }
    })()


    var jsonToCSV = function (dataSource, columns) {
      var csvData = '',
        values = []
      columns.forEach(function (item) {
        if (item.computed && typeof item.computed === 'function') {
          values.push(item.computed)
        } else if (item.dataIndex && typeof item.dataIndex === 'string') {
          values.push(item.dataIndex)
        } else {
          values.push(null)
        }
      })

      dataSource.forEach(function (item) {
        values.forEach(function (key) {
          if (typeof key === 'string') {
            csvData += item[key]
          } else if (typeof key === 'function') {
            csvData += key(item) || ''
          } else {
            csvData += ''
          }
          csvData += ','
        })
        csvData = csvData.slice(0, csvData.length - 1)
        csvData += '\r\n'
      })
      return csvData
    }
    var base64 = function (s) {
      return window.btoa(window.unescape(encodeURIComponent(s)))
    }
    var exportXLS = function (dataSource, columns, fileName) {
      var XLSData = 'data:application/vnd.ms-excel;base64,' + base64(jsonToXLS(dataSource, columns))
      download(XLSData, fileName)
    }
    var exportCSV = function (dataSource, columns, fileName) {
      var CSVData = 'data:application/csv;base64,' + base64(jsonToCSV(dataSource, columns))
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
      var arr = base64Data.split(','),
        mime = arr[0].match(/:(.*?);/)[1],
        bstr = atob(arr[1]),
        n = bstr.length,
        u8arr = new Uint8Array(n)
      while (n--) {
        u8arr[n] = bstr.charCodeAt(n)
      }
      return new Blob([u8arr], {
        type: mime
      })
    }
    return {
      exportXLS: exportXLS,
      exportCSV: exportCSV
    }
  }())
  if (typeof define === 'function' && define.amd) {
    define(function () {
      return JSON2EXCEL
    })
  } else if (typeof exports !== 'undefined') {
    if (typeof module !== 'undefined' && module.exports) {
      exports = module.exports = JSON2EXCEL
    }
    exports.JSON2EXCEL = JSON2EXCEL
  } else {
    global.JSON2EXCEL = JSON2EXCEL
  }
})(this)