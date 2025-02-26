// 工具函数
const fsPromises = require('fs').promises;
const nodeXlsx = require('node-xlsx'); // excel读取解析

exports.fetcher = async (url, params = {
  method: 'GET',
  headers: {},
  body: null
}, resolver = 'json') => {
  if (params.body !== null) {
    params.method = 'POST';
    params.body = JSON.stringify(params.body || '');
    const headers = params.headers || {};
    headers['Content-Type'] = 'application/json';
    params.headers = headers;
  }
  const response = await fetch(url, params);
  return response[resolver]();
};

exports.xlsxSaver = (filePath, name, data, cols) => fsPromises.writeFile(filePath, nodeXlsx.build([{
  name,
  data
}], {
  cellDates: false,
  sheetOptions: {
    '!cols': cols
  }
}));

// 图片内容数据转换器
exports.getTextFromPDF = async path => {
  const pdfjsLib = await import('pdfjs-dist');
  const doc = await pdfjsLib.getDocument(path).promise;
  const promises = [];
  for (let index = 0; index < doc.numPages; ++index) {
    promises.push(doc.getPage(index + 1).then(page => page.getTextContent()).then(content => {
      const { items } = content;
      const tokens = content.items.filter(item => item.str.trim()).map(item => item.str);
      return {
        items,
        tokens,
        text: tokens.join('')
      };
    }));
  }
  return Promise.all(promises);
}