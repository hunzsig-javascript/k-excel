require('script-loader!file-saver');
require('script-loader!xlsx/dist/xlsx.core.min');
require('script-loader!blob.js/Blob');
var JsZip = require("jszip");

/**
 * json 导出 excel
 * @date 2019/09/08
 * @author：450947795@qq.com
 * @developer: mzyhaohaoren@qq.com
 */
const changeData = function (data, filter) {
  let sj = data,
    f = filter,
    re = [];
  Array.isArray(data) ? (function () {
    //对象
    f ? (function () {
      //存在过滤
      sj.forEach(function (obj) {
        let one = [];
        filter.forEach(function (no) {
          one.push(obj[no]);
        });
        re.push(one);
      });
    })() : (function () {
      //不存在过滤
      sj.forEach(function (obj) {
        let col = Object.keys(obj);
        let one = [];
        col.forEach(function (no) {
          one.push(obj[no]);
        });
        re.push(one);
      });

    })();
  })() : (function () {
    re = sj;
  })();
  return re;
};


/**
 * 转换数据
 */
const sheetChangeData = function (data) {

  let ws = {};
  let range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };
  for (let R = 0; R !== data.length; ++R) {
    for (let C = 0; C !== data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      let cell = {
        v: data[R][C]
      };
      if (cell.v == null) continue;
      let cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';
      ws[cell_ref] = cell;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
};

const s2ab = function (s) {
  let buf = new ArrayBuffer(s.length);
  let view = new Uint8Array(buf);
  for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
};
const datenum = function (v, date1904) {
  if (date1904) v += 1462;
  let epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};
const fixdata = function (data) {
  let o = '',
    l = 0,
    w = 10240;
  for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
  o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
  return o;
};

const exportExcel = function () {

  const instance = {
    excelZip: function (data, dataLength, set) {
      const map = [];
      const sheetFilter = [];
      const sheetHeader = [];
      set.sheet.forEach(function (s) {
        sheetFilter.push(s.key);
        sheetHeader.push(s.value);
      });
      //根据需求是否分开打印表格与压缩zip
      const residue = dataLength.total % set.sheetLength;
      const pernum = parseInt(dataLength.total / set.sheetLength);
      const sumPage = residue ? pernum + 1 : pernum;
      let currentPage = 0;
      let dataOrder = 0;
      for (currentPage; currentPage <= sumPage - 1; currentPage += 1) {
        const temp = [];
        if (residue !== 0 && currentPage === sumPage - 1) {
          for (let i = 0; i <= residue - 1; i += 1) {
            const temp1 = {};
            for (let item = 0; item < sheetFilter.length; item += 1) {
              temp1[sheetFilter[item]] = eval('data[dataOrder].' + sheetFilter[item]);
            }
            dataOrder += 1;
            temp.push(temp1);
          }
        } else {
          for (let i = 0; i <= set.sheetLength - 1; i += 1) {
            const temp1 = {};
            for (let item = 0; item < sheetFilter.length; item += 1) {
              temp1[sheetFilter[item]] = eval('data[dataOrder].' + sheetFilter[item]);
            }
            dataOrder += 1;
            temp.push(temp1);
          }
        }
        map.push(temp);
      }
      let option = {};
      let datas = [];
      option.fileName = 'excel';
      map.forEach(function (mapitem, index) {
        datas.push({
          sheetName: 'sheet' + index,
          sheetFilter: sheetFilter,
          sheetHeader: sheetHeader,
          sheetData: mapitem
        });
      });
      option.datas = datas;
      instance.saveExcel(option);
    },
    saveExcel: function (options) {
      let _options = {
        fileName: options.fileName || 'download',
        datas: options.datas,
        workbook: {
          SheetNames: [],
          Sheets: {}
        }
      };
      let wb = _options.workbook;

      _options.datas.forEach(function (data, index) {
        let sheetHeader = data.sheetHeader || null;
        let sheetData = data.sheetData;
        let sheetName = data.sheetName || 'sheet' + (index + 1);
        let sheetFilter = data.sheetFilter || null;

        sheetData = changeData(sheetData, sheetFilter);

        if (sheetHeader) {
          sheetData.unshift(sheetHeader)
        }

        let ws = sheetChangeData(sheetData);

        ws['!merges'] = [];

        wb.SheetNames.push(sheetName);
        wb.Sheets[sheetName] = ws;

      });
      let wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
      });
      saveAs(new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
      }), _options.fileName + ".xlsx");

      /*   saveAs(new Blob([s2ab(wbout)], {
             type: "application/octet-stream"
         }), _options.fileName + ".xlsx");*/

      /* var zipExcel = s2ab(wbout);
       var zip = new JsZip();
       var excel = zip.folder('excel');
       excel.file(sheetName, zipExcel, {base64: true});

       zip.generateAsync({ type: 'blob' }).then(function(content) {
         // see FileSaver.js
         saveAs(content, 'excel.zip');
       });*/
    },
    excelPull: function (element, set, then) {
      let wb = '';
      let rABS = false;
      let suffix = element.name.split('.')[1];
      if (suffix !== 'xls' && suffix !== 'xlsx') {
        alert('导入的文件格式不正确!');
        return;
      }
      const f = element;
      let reader = new FileReader();
      reader.onload = function (e) {
        let data = e.target.result, rowBox = [], excelPullBox = [];
        if (rABS) {
          wb = XLSX.read(btoa(fixdata(data)), {
            type: 'base64',
          });
        } else {
          wb = XLSX.read(data, {
            type: 'binary',
          });
        }
        const sheet1 = wb.Sheets[wb.SheetNames[0]];
        const range = XLSX.utils.decode_range(sheet1['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
          let row = [];
          for (let C = range.s.c; C <= range.e.c; ++C) {
            let row_value = null;
            let cell_address = {c: C, r: R};
            let cell = XLSX.utils.encode_cell(cell_address);
            if (sheet1[cell]) {
              row_value = sheet1[cell].v;
            } else {
              row_value = '';
            }
            row.push(row_value);
          }
          rowBox.push(row);
        }
        rowBox.forEach(function (i) {
          let excelPull = {};
          set.forEach(function (item, index) {
            excelPull[set[index].key] = i[index];
          });
          excelPullBox.push(excelPull);
        });
        then(excelPullBox);
      };
      if (rABS) {
        reader.readAsArrayBuffer(f);
      } else {
        reader.readAsBinaryString(f);
      }
    },
  };

  return instance;
};

module.exports = exportExcel;
