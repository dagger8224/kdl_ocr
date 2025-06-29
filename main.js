// entry

const fs = require('fs');
const path = require('path');
const { app, BrowserWindow, ipcMain, shell } = require('electron');
const utils = require('./utils.js');

const loginErrorMessage = '账号或者密码有误，请重新输入';

const profile = {
  config: {},
  licenseInfo: {
    loginTime: 0
  }
};

const dataPath = '../data';
const inputFilePath = `${ dataPath }/input.pdf`;
const outputFilePath = `${ dataPath }/output.xlsx`;
const productTypeFilePath = `${ dataPath }/productTypes.xlsx`;
const priceFilePath = `${ dataPath }/prices.xlsx`;
const dealerFilePath = `${ dataPath }/dealers.xlsx`;
const orderFilePath = `${ dataPath }/orders.xlsx`;
const splitOrderFilePath = `${ dataPath }/splitOrders.xlsx`;
const invalidOrderFilePath = `${ dataPath }/invalidOrders.xlsx`;
const billFilePath = `${ dataPath }/bill.xlsx`;
const receiptFilePath = `${ dataPath }/receipt.xlsx`;
const mergeResultFilePath1 = `${ dataPath }/mergeResult1.xlsx`;
const mergeResultFilePath2 = `${ dataPath }/mergeResult2.xlsx`;

// 观察文件状态
const fileExistsStateWatcher = () => BrowserWindow.getAllWindows()[0].webContents.send('fileExistsState', {
  openPdfFileButton: fs.existsSync(inputFilePath),
  openExcelFileButton: fs.existsSync(outputFilePath),
  openProductTypeFileButton: fs.existsSync(productTypeFilePath),
  openPriceFileButton: fs.existsSync(priceFilePath),
  openDealerFileButton: fs.existsSync(dealerFilePath),
  openOrderFileButton: fs.existsSync(orderFilePath),
  openSplitOrderFileButton: fs.existsSync(splitOrderFilePath),
  openInvalidOrderFileButton: fs.existsSync(invalidOrderFilePath),
  openBillFileButton: fs.existsSync(billFilePath),
  openReceiptFileButton: fs.existsSync(receiptFilePath),
  openMergeResultFileButton1: fs.existsSync(mergeResultFilePath1),
  openMergeResultFileButton2: fs.existsSync(mergeResultFilePath2),
  importOrderFileButton: fs.existsSync(productTypeFilePath) && fs.existsSync(priceFilePath) && fs.existsSync(dealerFilePath)
});

fs.existsSync(dataPath) || fs.mkdirSync(dataPath, 744);
fs.watch(dataPath, fileExistsStateWatcher);

// 创建主窗口

const createWindow = async () => {
  const mainWindow = new BrowserWindow({
    width: 1920,
    height: 1080,
    resizable: true,
    webPreferences: {
      devTools: true,
      nodeIntegration: true,
      // webSecurity: false,
      preload: path.join(__dirname, 'preload.js'),
      // requestedExecutionLevel: 'requireAdministrator' // 或者 'highestAvailable'
    },
  });
  mainWindow.removeMenu();
  // 打开开发工具
  // mainWindow.webContents.openDevTools();
  mainWindow.loadFile('page/index.html');
  mainWindow.maximize();
  // mainWindow.loadURL('https://towan-cos-beijing-1304741629.cos.ap-beijing.myqcloud.com/batchProcessApp/index.html');
  mainWindow.webContents.on('dom-ready', fileExistsStateWatcher);
};

// 这段程序将会在 Electron 结束初始化
// 和创建浏览器窗口的时候调用
// 部分 API 在 ready 事件触发后才能使用。
app.whenReady().then(() => {
  createWindow();
  // 在 macOS 系统内, 如果没有已开启的应用窗口
  // 点击托盘图标时通常会重新创建一个新窗口
  app.on('activate', () => {
    if (!BrowserWindow.getAllWindows().length) {
      createWindow();
    }
  });
});

// 除了 macOS 外，当所有窗口都被关闭的时候退出程序。因此，通常
// 对应用程序和它们的菜单栏来说应该时刻保持激活状态，直到用户使用 Cmd + Q 明确退出
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
    console.log('app is closed');
  }
});

/* 前端事件消息处理 */

// 获取配置信息
ipcMain.handle('getVersion', async () => {
  profile.config = await utils.fetcher('https://towan-cos-beijing-1304741629.cos.ap-beijing.myqcloud.com/deepSeek/config.json');
  return profile.config.version;
});

// 获取授权信息
ipcMain.handle('getLoginInfo', async () => {
  return profile.licenseInfo;
});

// 验证授权信息
ipcMain.handle('verifyLicense', async (_, account, password) => {
  const licenseMap = await utils.fetcher('https://towan-cos-beijing-1304741629.cos.ap-beijing.myqcloud.com/deepSeek/license.json');
  const licenseInfo = licenseMap[account];
  if (licenseInfo) {
    if (password !== licenseInfo.password) {
      return {
        message: loginErrorMessage
      };
    }
    const { username, startTime, authMode } = licenseInfo;
    profile.licenseInfo = {
      username,
      authMode: authMode === 3 ? ' (试用版)' : '',
      leftDays: Math.max(authMode - Math.floor((Date.now() - new Date(startTime).getTime()) / (24 * 60 * 60 * 1000)), 0),
      loginTime: Date.now()
    };
    BrowserWindow.getAllWindows()[0].webContents.send('updateHeader', profile.licenseInfo);
    return {
      message: '登录成功',
      licenseInfo: profile.licenseInfo
    };
  } else {
    return {
      message: loginErrorMessage
    };
  }
});

// 退出登录
ipcMain.handle('logout', () => {
  profile.licenseInfo = {
    loginTime: 0
  };
  return true;
});

const findToken = (tokens, prefixes) => {
  tokens = tokens.map(token => token.trim());
  const prefix = prefixes.find(prefix => tokens.find(token => token.startsWith(prefix)));
  if (prefix) {
    const token = tokens.find(token => token.startsWith(prefix));
    return token.replace(prefix, '').trim() || '-';
  }
  return '-';
};

/* ============ 工具函数 ============ */
// token 再按空白、斜杠切碎 → 去括号、空串
function atomizeToken(tok) {
  return tok
    .split(/\s+/)
    .flatMap(s => s.split('/'))
    .map(s => s.replace(/[()]/g, '').trim())
    .filter(Boolean);
}
// 8 位日期 → yyyy/mm/dd
const fmtDate8 = s =>
  /^\d{8}$/.test(s) ? s.replace(/(\d{4})(\d{2})(\d{2})/, '$1/$2/$3') : s;
// 温度归一化 "低～高℃"
function normalizeTemp(str = '-') {
  // 处理 +02°Cto08°C 这样的格式
  const specialFormat = str.match(/\+?(\d+)°C.*?(\d+)°C/i);
  if (specialFormat) {
    return `${parseInt(specialFormat[1], 10)}～${parseInt(specialFormat[2], 10)}℃`;
  }
  
  // 原有格式处理
  const m = str.match(/(\d+).*?(\d+).*?℃/);
  return m ? `${parseInt(m[1], 10)}～${parseInt(m[2], 10)}℃` : str;
}

/* ============ 主处理器 ============ */
ipcMain.handle('importPdfFile', async (_, filePath, isType1) => {
  try {
    /* ① 读 PDF，生成 tokens 与全文字符串 */
    fs.copyFileSync(filePath, inputFilePath, fs.constants.COPYFILE_FICLONE);
    const pages = await utils.getTextFromPDF(inputFilePath);

    let tokens = [];
    let pdfText = '';
    pages.forEach(p => {
      tokens.push(...p.tokens.flatMap(atomizeToken));
      pdfText += ' ' + (p.text ? p.text : p.tokens.join(' '));
    });

    /* ② 清分页 token */
    const isPageTok = s =>
      /^Page\s*\d+\s*of\s*\d+$/i.test(s) ||
      /^第\d+页共\d+页$/.test(s) || /^第\d+页$/.test(s) ||
      /^共\d+页$/.test(s) || /^\d+页$/.test(s);
    tokens = tokens.filter(t => !isPageTok(t));

    /* ③ 删除 “To … 有效期限至” */
    let a, b;
    do {
      a = tokens.indexOf('To');
      b = tokens.findIndex(t => t.startsWith('有效期限至'));
      if (a > -1 && b > -1 && b > a) tokens.splice(a, b - a + 1);
    } while (a > -1 && b > -1);

    /* ④ 删除尾部声明 */
    const tail = tokens.findIndex(t =>
      /(These items are controlled|RESERVATION CLAUSE)/i.test(t)
    );
    if (tail !== -1) tokens = tokens.slice(0, tail);

    /* ⑤ 公共 4 字段（正则 & 双重兜底） */
    const dateM = pdfText.match(/通知日期[:：]?\s*([0-9]{2}\/[0-9]{2}\/[0-9]{4})/i);
    const noteM = pdfText.match(/发货通知单号[:：]?\s*(SI[0-9A-Z]{8,})/i);
    const custM = pdfText.match(/(LS-[0-9A-Z\-]+)/i);

    // —— robust Siemens order —— //
    let siemensOrdNo = '';
    
    // 尝试多种匹配模式
    // 1. 直接匹配"西门子订单号"后面的内容
    const siemensOrderMatch = pdfText.match(/西门子订单号[:：]?\s*([\w\d\-]{5,15})/i);
    if (siemensOrderMatch) {
      siemensOrdNo = siemensOrderMatch[1];
    } else {
      // 2. 查找包含"订单号"和"西门子"的上下文
      const orderContext = pdfText.match(/订单号[:：]?\s*([\w\d\-]{5,15}).*?西门子|西门子.*?订单号[:：]?\s*([\w\d\-]{5,15})/i);
      if (orderContext) {
        siemensOrdNo = orderContext[1] || orderContext[2];
      } else {
        // 3. 原有Siemens关键词查找逻辑，但放宽数字格式要求
        const siIdx = pdfText.search(/Siemens|西门子/i);
        if (siIdx !== -1) {
          const snip = pdfText.slice(siIdx, siIdx + 200);
          // 放宽匹配条件，查找5-15位的数字字母组合
          const dig = snip.match(/([\d\w\-]{5,15})/); 
          if (dig) siemensOrdNo = dig[1];
        }
        
        // 4. 在tokens中查找可能的订单号
        if (!siemensOrdNo) {
          const pos = tokens.findIndex(t => /Siemens|西门子|订单/i.test(t));
          if (pos !== -1) {
            for (let j = pos + 1; j < pos + 30 && j < tokens.length; j++) {
              // 放宽匹配条件
              if (/^[\d\w\-]{5,15}$/.test(tokens[j]) && !/^(UEG:|REF:)/.test(tokens[j])) {
                siemensOrdNo = tokens[j];
                break;
              }
            }
          }
        }
      }
    }

    const deliveryDate   = dateM ? dateM[1] : '';
    const deliveryNoteNo = noteM ? noteM[1] : '';
    const customerOrdNo  = custM ? custM[1] : '';

    const baseRow = isType1
      ? [deliveryDate, deliveryNoteNo, customerOrdNo, siemensOrdNo]
      : [deliveryDate, customerOrdNo, siemensOrdNo];

    /* ⑥ 表头 */
    const header = isType1
      ? ['发货日期','发货单号','客户订单号','西门子订单号',
         '西门子编码','储存温度','UEG','批号','数量','单位','效期','生产日期']
      : ['发货日期','客户订单号','西门子订单号','西门子编码',
         '储存温度','UEG','数量','单价','金额'];
    const out = [header];

    /* ⑦ 正则工具 */
    const seq6  = /^\d{6}$/;                 // 行序号
    const code8 = /^\d{8}$/;                 // 西门子编码
    const tempR = /(\d+[~～\-至to]\d+℃)|(\+\d+°C.*?\d+°C)/i;
    const uegR  = /^(UEG:|REF:)\d+$/i;
    const qtyR  = /^\d+(\.\d+)?$/;
    const unitR = /(Piece|件|PCE|pack|盒)/i;
    const mmddyyyy = /^\d{2}\/\d{2}\/\d{4}$/;

    /* ⑧ 解析每个序号块 */
    let i = 0;
    while (i < tokens.length) {
      if (!seq6.test(tokens[i])) { i++; continue; }
      const start = i++;
      while (i < tokens.length && !seq6.test(tokens[i])) i++;
      const block = tokens.slice(start, i);

      const code = block.find(t => code8.test(t)) || '';
      const rawTemp = (block.find(t => tempR.test(t))
                      || block.join('').match(tempR)?.[0]) || '-';
      const temp = normalizeTemp(rawTemp);

      const uTok = block.find(t => uegR.test(t)) || '-';
      const uVal = uTok.replace(/^(UEG:|REF:)/i, '');

      const prefix = [...baseRow, code, temp, uVal];

      if (isType1) {
        // 滑窗找批号五元组
        for (let k = 0; k + 4 < block.length; k++) {
          const [batch, qty, unit, exp, mfg] = block.slice(k, k + 5);
          if (
            qtyR.test(qty) && unitR.test(unit) &&
            (mmddyyyy.test(exp) || /^\d{8}$/.test(exp)) &&
            (mmddyyyy.test(mfg) || /^\d{8}$/.test(mfg))
          ) {
            out.push([
              ...prefix,
              batch,
              qty,
              unit,
              fmtDate8(exp),
              fmtDate8(mfg)
            ]);
            k += 4;
          }
        }
      } else {
        // 金额版
        const qty   = block.find(t => qtyR.test(t)) || '-';
        const price = block.find((t, idx) => qtyR.test(t) && idx > block.indexOf(qty)) || '-';
        const amt   = [...block].reverse().find(qtyR.test.bind(qtyR)) || '-';
        out.push([...prefix, qty, price, amt]);
      }
    }

    /* ⑨ 输出 Excel */
    utils.xlsxSaver(
      outputFilePath,
      [{ name: '识别结果', data: out }],
      header.map(() => ({ wch: 18 }))
    );
    return '内容识别完成';

  } catch (err) {
    return err.code === 'EBUSY'
      ? '请关闭已打开的 pdf / 结果文件后再试'
      : `操作失败：${err.message}`;
  }
});

/*
// 导入pdf数据
ipcMain.handle('importPdfFile', async (_, filePath, isType1) => {
  try {
    fs.copyFileSync(filePath, inputFilePath, fs.constants.COPYFILE_FICLONE);
    const contentList = await utils.getTextFromPDF(inputFilePath);
    const pageCount = contentList.length;
    let fullTokens = contentList.map(page => page.tokens).flat();
    const SiemensOrderIndex = fullTokens.findIndex(token => token.startsWith('Siemens Reference Number'));
    const baseRowData = isType1 ? [fullTokens[SiemensOrderIndex + 2], fullTokens[fullTokens.indexOf('发货通知单号:') + 1], fullTokens[SiemensOrderIndex + 1], fullTokens[SiemensOrderIndex + 3]] : [fullTokens[SiemensOrderIndex + 3], fullTokens[SiemensOrderIndex + 2], fullTokens[SiemensOrderIndex + 4]];
    let startIndex = 0;
    let endIndex = 0;
    do {
      startIndex = fullTokens.indexOf('To');
      endIndex = fullTokens.findIndex(token => token.startsWith('有效期限至'));
      fullTokens.splice(startIndex, endIndex - startIndex + 1);
    } while (startIndex !== -1 && endIndex !== -1);
    let pageIndex = 1;
    while (pageIndex <= pageCount) {
      startIndex = fullTokens.findIndex(token => token.startsWith(`Page ${ pageIndex } of`));
      endIndex = fullTokens.findIndex(token => token.endsWith(isType1 ? `${ pageCount }页` : `${ pageCount } 页`));
      fullTokens.splice(startIndex, endIndex - startIndex + 1);
      pageIndex++;
    }
    fullTokens = fullTokens.slice(0, fullTokens.findIndex(token => token.startsWith(isType1 ? 'These items are controlled by the U.S.' : 'RESERVATION CLAUSE')));
    const headerRow = isType1 ? ['发货日期', '发货单号', '客户订单号', '西门子订单号', '西门子编码', '储存温度', 'UEG'] : ['发货日期', '客户订单号', '西门子订单号', '西门子编码', '储存温度', 'UEG', '数量', '单价', '金额'];
    const resultData = [headerRow];
    const groupPrefix = ['批号', '数量', '单位', '效期', '生产日期']; // 可能有多组批号效期数据
    const groupLength = groupPrefix.length;
    let baseIndex = isType1 ? 0 : 11;
    let itemStartNumber = fullTokens[baseIndex] - 0; // 起始序号转为数字，应该为10/1000
    let itemStartNumberIndex = 0;
    while (fullTokens.length) {
      const rowData = [...baseRowData];
      rowData.push(fullTokens[baseIndex + 1]); // 西门子编码
      const isFirst = itemStartNumber === 10;
      itemStartNumber += isType1 ? 10 : 1000;
      itemStartNumberIndex = fullTokens.indexOf((itemStartNumber + '').padStart(6, '0'));
      let maxCount = 0;
      while (itemStartNumberIndex === -1 && maxCount < 100) {
        maxCount++;
        itemStartNumber += isType1 ? 10 : 1000;
        itemStartNumberIndex = fullTokens.indexOf((itemStartNumber + '').padStart(6, '0'));
      }
      itemStartNumberIndex === -1 && (itemStartNumberIndex = fullTokens.length);
      const currentTokens = fullTokens.slice(0, itemStartNumberIndex);
      rowData.push(findToken(currentTokens, ['Storage Condition:', 'Temperature Condition: D3 Description:', 'Temperature Condition: D4 Description:']));
      rowData.push(findToken(currentTokens, ['Siemens Sort No: UEG:', 'REF:']));
      if (isType1) {
        // 第一页结构与后面的不一样，分别处理
        startIndex = 0;
        if (isFirst) {
          startIndex = currentTokens.findIndex(token => token.trim().endsWith('Manuf. Date'));
        } else {
          startIndex = currentTokens.findIndex(token => token.trim().startsWith('CFDA License:'));
          if (startIndex === -1) {
            startIndex = currentTokens.findIndex(token => token.trim().startsWith('Country of Origin:'));
          }
        }
        const splitTokens = currentTokens.slice(startIndex + 1, itemStartNumberIndex).join('/').split('/').filter(token => token.trim()).map(token => token.trim());
        const group = [];
        splitTokens.forEach((token, index) => {
          const resolvedIndex = index % groupLength;
          const fieldName = groupPrefix[resolvedIndex];
          headerRow.includes(fieldName) || headerRow.push(fieldName);
          if ((resolvedIndex === 3 || resolvedIndex === 4) && !token.includes('/')) {
            token = `${ token.slice(0, 4) }/${ token.slice(4, 6) }/${ token.slice(6, 8) }`;
          }
          const rowIndex = Math.floor(index / groupLength);
          const row = group[rowIndex] || [];
          row.push(token);
          group[rowIndex] = row;
        });
        group.forEach(row => resultData.push([...rowData, ...row]));
        if (!group.length) {
          resultData.push(rowData);
        }
      } else {
        rowData.push(fullTokens[baseIndex + 3]); // 数量
        rowData.push(fullTokens[baseIndex + 4]); // 单价
        rowData.push(fullTokens[baseIndex + 5]); // 金额
        resultData.push(rowData);
      }
      fullTokens = fullTokens.slice(itemStartNumberIndex);
      baseIndex = 0;
    }
    utils.xlsxSaver(outputFilePath, [{
      name: '识别结果',
      data: resultData
    }], headerRow.map(() => ({ wch: 16 })));
    return '内容识别完成';
  } catch (error) {
    return error.code === 'EBUSY' ? `请关闭已打开的pdf和转换结果文件后再试` : `操作失败：${ error.message }`;
  }
});
*/

const SiemensCodeField = '西门子编码';
const productTypeField = '货品分类';
const operatorField = '操作人';
const timeField = '导入时间';
// 导入产品类别表
ipcMain.handle('importProductTypeFile', async (_, filePath) => {
  try {
    const productTypeData = utils.xlsxParser(filePath);
    const productTypeDataHeaderRow = productTypeData[0];
    if (!productTypeDataHeaderRow?.includes(SiemensCodeField)) {
      return '产品类别表中没有西门子编码数据';
    }
    if (!productTypeDataHeaderRow?.includes(productTypeField)) {
      return '产品类别表中没有货品分类数据';
    }
    const resolvedProductTypeData = {};
    utils.dataResolver(productTypeData).forEach(row => {
      if (row[SiemensCodeField] && row[productTypeField]) {
        resolvedProductTypeData[row[SiemensCodeField]] = row;
      }
    });
    const operator = profile.licenseInfo.username;
    const currentTime = new Date().toLocaleString();
    let resolvedData = [];
    if (fs.existsSync(productTypeFilePath)) {
      const resolvedExistProductTypeData = {};
      utils.dataResolver(utils.xlsxParser(productTypeFilePath)).forEach(row => (resolvedExistProductTypeData[row[SiemensCodeField]] = row));
      Object.keys(resolvedProductTypeData).forEach(SiemensCode => {
        const row = resolvedProductTypeData[SiemensCode];
        const existRow = resolvedExistProductTypeData[SiemensCode] || {};
        existRow[SiemensCodeField] = SiemensCode;
        existRow[productTypeField] = row[productTypeField];
        resolvedExistProductTypeData[SiemensCode] = existRow;
      });
      resolvedData = Object.keys(resolvedExistProductTypeData).map(SiemensCode => {
        const row = resolvedExistProductTypeData[SiemensCode];
        return [row[SiemensCodeField], row[productTypeField], operator, currentTime];
      });
    } else {
      resolvedData = Object.values(resolvedProductTypeData).map(row => [row[SiemensCodeField], row[productTypeField], operator, currentTime]);
    }
    resolvedData.unshift([SiemensCodeField, productTypeField, operatorField, timeField]);
    await utils.xlsxSaver(productTypeFilePath, [{
      name: '产品类别表',
      data: resolvedData
    }], [{ wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 25 }]);
    return '导入产品类别表完成';
  } catch (error) {
    return `导入产品类别表失败：${ error.message }`;
  }
});

const productCodeField = '产品编码';
const standardSellPriceField = '标准价格-销售价格';
const standardBuyPriceField = '标准价格-采购价格';
const BeijingSpecialSellPriceField = '北京特殊价-销售价格';
const BeijingSpecialBuyPriceField = '北京特殊价-采购价格';
const IAVBPSellPriceField = 'IA VBP-销售价格';
const IAVBPBuyPriceField = 'IA VBP-采购价格';
const VBP1SellPriceField = 'VBP1-销售价格';
const VBP1BuyPriceField = 'VBP1-采购价格';
const VBP2SellPriceField = 'VBP2-销售价格';
const VBP2BuyPriceField = 'VBP2-采购价格';
const resolvedPriceDataHeaderRow = [SiemensCodeField, productCodeField, standardSellPriceField, standardBuyPriceField, BeijingSpecialSellPriceField, BeijingSpecialBuyPriceField, IAVBPSellPriceField, IAVBPBuyPriceField, VBP1SellPriceField, VBP1BuyPriceField, VBP2SellPriceField, VBP2BuyPriceField];
// 导入产品价格表
ipcMain.handle('importPriceFile', async (_, filePath) => {
  try {
    const priceData = utils.xlsxParser(filePath);
    const priceDataHeaderRow = priceData[0];
    for (let index = 0; index < resolvedPriceDataHeaderRow.length; ++index) {
      if (!priceDataHeaderRow?.includes(resolvedPriceDataHeaderRow[index])) {
        return `产品价格表中没有${ resolvedPriceDataHeaderRow[index] }数据`;
      }
    }
    const resolvedPriceData = {};
    utils.dataResolver(priceData).forEach(row => {
      if (row[SiemensCodeField]) {
        resolvedPriceData[row[SiemensCodeField]] = row;
      }
    });
    const operator = profile.licenseInfo.username;
    const currentTime = new Date().toLocaleString();
    let resolvedData = [];
    if (fs.existsSync(priceFilePath)) {
      const resolvedExistPriceData = {};
      utils.dataResolver(utils.xlsxParser(priceFilePath)).forEach(row => (resolvedExistPriceData[row[SiemensCodeField]] = row));
      Object.keys(resolvedPriceData).forEach(SiemensCode => {
        const row = resolvedPriceData[SiemensCode];
        const existRow = resolvedExistPriceData[SiemensCode] || {};
        resolvedPriceDataHeaderRow.forEach(field => {
          existRow[field] = row[field];
        });
        resolvedExistPriceData[SiemensCode] = existRow;
      });
      resolvedData = Object.keys(resolvedExistPriceData).map(SiemensCode => {
        const row = resolvedExistPriceData[SiemensCode];
        return [...resolvedPriceDataHeaderRow.map(field => row[field]), operator, currentTime];
      });
    } else {
      resolvedData = Object.values(resolvedPriceData).map(row => [...resolvedPriceDataHeaderRow.map(field => row[field]), operator, currentTime]);
    }
    resolvedData.unshift([...resolvedPriceDataHeaderRow, operatorField, timeField]);
    await utils.xlsxSaver(priceFilePath, [{
      name: '产品价格表',
      data: resolvedData
    }], [{ wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }]);
    return '导入产品价格表完成';
  } catch (error) {
    return `导入产品价格表失败：${ error.message }`;
  }
});

const dealerCodeField = '客户编码';
const dealerNameField = '客户名称';
const areaField = '区域';
const priceLevelField = '价格级别';
const ShiptoPartyField = 'ShiptoParty';
const resolvedDealerDataHeaderRow = [dealerCodeField, dealerNameField, areaField, priceLevelField, ShiptoPartyField];
// 导入经销商信息表
ipcMain.handle('importDealerFile', async (_, filePath) => {
  try {
    const dealerData = utils.xlsxParser(filePath);
    const dealerDataHeaderRow = dealerData[0];
    if (!dealerDataHeaderRow?.includes(dealerCodeField)) {
      return '经销商信息表中没有客户编码数据';
    }
    if (!dealerDataHeaderRow?.includes(priceLevelField)) {
      return '经销商信息表中没有价格级别数据';
    }
    if (!dealerDataHeaderRow?.includes(ShiptoPartyField)) {
      return '经销商信息表中没有子账号数据';
    }
    const resolvedDealerData = [];
    utils.dataResolver(dealerData).forEach(row => {
      if (row[dealerCodeField] && row[ShiptoPartyField]) {
        resolvedDealerData.push(row);
      }
    });
    const operator = profile.licenseInfo.username;
    const currentTime = new Date().toLocaleString();
    const resolvedData = (fs.existsSync(dealerFilePath) ? [...utils.dataResolver(utils.xlsxParser(dealerFilePath)), ...resolvedDealerData] : resolvedDealerData).map(row => [...resolvedDealerDataHeaderRow.map(field => row[field]), operator, currentTime]);
    resolvedData.unshift([...resolvedDealerDataHeaderRow, operatorField, timeField]);
    await utils.xlsxSaver(dealerFilePath, [{
      name: '经销商信息表',
      data: resolvedData
    }], [{ wch: 20 }, { wch: 35 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 20 }, { wch: 20 }]);
    return '导入经销商信息表完成';
  } catch (error) {
    return `导入经销商信息表失败：${ error.message }`;
  }
});

const orderField = '订单';
const subOrderField = '子账号';
const orderTypeField = '订单类型';
const productCountField = '要货数量';
const rateField = '毛利率';
const totalBuyAmountField = '采购总价';
const amountField = '欠货金额';
const invalidReasonField = '无效原因';
const sellPriceField = '销售价格';
const buyPriceField = '采购价格';
const requiredOrderDataHeaderRow = [dealerCodeField, SiemensCodeField, productCountField, amountField];
const commonSheetHeaderRow = [orderTypeField, productTypeField, productCountField, sellPriceField, priceLevelField, buyPriceField, rateField, totalBuyAmountField];
const proteinSheetHeaderRow = [dealerCodeField, dealerNameField, SiemensCodeField, ...commonSheetHeaderRow];
const otherSheetHeaderRow = [dealerCodeField, dealerNameField, SiemensCodeField, orderField, subOrderField, ...commonSheetHeaderRow];
const invalidOrderHeaderRow = [dealerCodeField, dealerNameField, SiemensCodeField, invalidReasonField];
const priceStrategy = [{
  sellPriceField: standardSellPriceField,
  buyPriceField: standardBuyPriceField,
  priceLevelField: '标准价格'
}, {
  sellPriceField: BeijingSpecialSellPriceField,
  buyPriceField: BeijingSpecialBuyPriceField,
  priceLevelField: '北京特殊价'
}, {
  sellPriceField: IAVBPSellPriceField,
  buyPriceField: IAVBPBuyPriceField,
  priceLevelField: 'IA VBP'
}, {
  sellPriceField: VBP1SellPriceField,
  buyPriceField: VBP1BuyPriceField,
  priceLevelField: 'VBP1'
}, {
  sellPriceField: VBP2SellPriceField,
  buyPriceField: VBP2BuyPriceField,
  priceLevelField: 'VBP2'
}];
// 导入原始订单数据表并拆分
ipcMain.handle('importOrderFile', async (_, filePath) => {
  try {
    fs.copyFileSync(filePath, orderFilePath, fs.constants.COPYFILE_FICLONE);
    const orderDataSheets = utils.xlsxSheetParser(orderFilePath, {
      cellDates: true
    });
    const orderData = [];
    for (let sheetIndex = 0; sheetIndex < orderDataSheets.length; ++sheetIndex) {
      const sheet = orderDataSheets[sheetIndex];
      const orderType = sheet.name;
      if (['临采订单', 'PTO订单'].includes(orderType)) {
        const data = sheet.data || [];
        const orderDataHeaderRow = data[0] || [];
        for (let index = 0; index < requiredOrderDataHeaderRow.length; ++index) {
          if (!orderDataHeaderRow?.includes(requiredOrderDataHeaderRow[index])) {
            return `${ orderType }中没有${ requiredOrderDataHeaderRow[index] }数据`;
          }
        }
        utils.dataResolver(data).forEach(row => {
          row[orderTypeField] = orderType;
          orderData.push(row);
        });
      }
    }
    if (!orderData.length) {
      return '原始订单数据表中没有名为“临采订单”或者“PTO订单”的sheet页';
    }
    const resolvedProductTypeData = {};
    utils.dataResolver(utils.xlsxParser(productTypeFilePath)).forEach(row => {
      resolvedProductTypeData[row[SiemensCodeField]] = row;
    });
    const resolvedPriceData = {};
    utils.dataResolver(utils.xlsxParser(priceFilePath)).forEach(row => {
      resolvedPriceData[row[SiemensCodeField]] = row;
    });
    const resolvedDealerData = {};
    utils.dataResolver(utils.xlsxParser(dealerFilePath)).forEach(row => {
      const dealerCode = row[dealerCodeField];
      if (resolvedDealerData[dealerCode]) {
        resolvedDealerData[dealerCode].push(row);
      } else {
        resolvedDealerData[dealerCode] = [row];
      }
    });
    const splitOrders = {}; // 根据订单类型/货品分类/子账号信息合并订单
    const invalidOrders = [];
    const proteinSheet = {
      name: '蛋白',
      data: []
    };
    const otherSheet = {
      name: 'LS和配件',
      data: []
    };
    orderData.forEach(row => {
      const SiemensCode = row[SiemensCodeField];
      const productType = resolvedProductTypeData[SiemensCode]?.[productTypeField];
      if (productType) {
        if (!['LS', '配件', '蛋白'].includes(productType)) {
          row[invalidReasonField] = `订单货品分类“${ productType }”不受支持`;
          return invalidOrders.push(row);
        } else {
          row[productTypeField] = productType;
          const resolvedProductCount = Number(row[productCountField]);
          const sellPrice = Math.floor(Number(row[amountField]) / resolvedProductCount);
          const priceRow = resolvedPriceData[SiemensCode];
          if (!priceRow) {
            row[invalidReasonField] = '产品价格表中没有找到与当前西门子编码匹配的数据';
            return invalidOrders.push(row);
          }
          const strategy = priceStrategy.find(strategy => priceRow[strategy.sellPriceField] === sellPrice);
          if (strategy) {
            row[sellPriceField] = sellPrice;
            row[buyPriceField] = priceRow[strategy.buyPriceField];
            row[priceLevelField] = strategy.priceLevelField;
            const dealerCode = row[dealerCodeField];
            const dealerList = resolvedDealerData[dealerCode];
            if (dealerList) {
              const dealerRow = row[priceLevelField] === 'VBP1' ? dealerList.find(dealer => dealer[priceLevelField] === 'VBP1') || dealerList[0] : dealerList.find(dealer => dealer[priceLevelField] !== 'VBP1');
              if (dealerRow) {
                row[subOrderField] = dealerRow[ShiptoPartyField];
              } else {
                row[invalidReasonField] = `没有找到与价格级别“${ row[priceLevelField] }”相匹配的子账号数据`;
                return invalidOrders.push(row);
              }
            } else {
              row[invalidReasonField] = `经销商信息表中没有找到与客户编码“${ dealerCode }”相匹配的经销商数据`;
              return invalidOrders.push(row);
            }
          } else {
            row[invalidReasonField] = `产品价格表中没有找到与订单销售价格“${ sellPrice }”相匹配的产品数据`;
            return invalidOrders.push(row);
          }
          const resolvedSellPrice = Number(row[sellPriceField]);
          row[rateField] = `${ (100 - Number(row[buyPriceField]) * 100 / resolvedSellPrice).toFixed(2) }%`;
          row[totalBuyAmountField] = resolvedProductCount * resolvedSellPrice;
          if (productType === '蛋白') {
            proteinSheet.data.push(proteinSheetHeaderRow.map(field => row[field]));
          } else {
            const uniqueKey = `${ productType }-${ row[orderTypeField] }-${ row[subOrderField] }`; // 由货品分类 - 订单类型 - 子账号生成唯一键
            if (splitOrders[uniqueKey]) {
              splitOrders[uniqueKey].push(row);
            } else {
              splitOrders[uniqueKey] = [row];
            }
          }
        }
      } else {
        row[invalidReasonField] = '产品类别表中没有找到与当前西门子编号对应的货品分类数据';
        return invalidOrders.push(row);
      }
    });
    proteinSheet.data.unshift(proteinSheetHeaderRow);
    otherSheet.data.unshift(otherSheetHeaderRow);
    Object.values(splitOrders).forEach((orders, index) => orders.forEach(order => {
      order[orderField] = `${ orderField }${ index + 1 }`;
      otherSheet.data.push(otherSheetHeaderRow.map(field => order[field]));
    }));
    const wchList = otherSheetHeaderRow.map(_ => ({ wch: 15 }));
    wchList[1].wch = 30;
    await utils.xlsxSaver(splitOrderFilePath, [otherSheet, proteinSheet], wchList);
    const invalidOrderData = invalidOrders.map(order => invalidOrderHeaderRow.map(field => order[field]));
    invalidOrderData.unshift(invalidOrderHeaderRow);
    await utils.xlsxSaver(invalidOrderFilePath, [{
      name: '无效订单',
      data: invalidOrderData
    }], [{ wch: 15 }, { wch: 30 }, { wch: 15 }, { wch: 60 }]);
    return '拆分订单完成';
  } catch (error) {
    return `拆分订单失败：${ error.message }`;
  }
});

// 导入发票列表文件
ipcMain.handle('importBillFile', (_, filePath) => {
  try {
    fs.copyFileSync(filePath, billFilePath, fs.constants.COPYFILE_FICLONE);
    return '导入发票列表数据完成';
  } catch (error) {
    return `导入发票列表数据失败：${ error.message }`;
  }
});

// 导入收款列表文件
ipcMain.handle('importReceiptFile', (_, filePath) => {
  try {
    fs.copyFileSync(filePath, receiptFilePath, fs.constants.COPYFILE_FICLONE);
    return '导入收款列表数据完成';
  } catch (error) {
    return `导入收款列表数据失败：${ error.message }`;
  }
});

// 聚合发票数据
const billCodeField = '发票号码';
const billAmountField = '发票金额';
const billDateField = '发票日期';
const receiptAmountField = '金额';
const receiptDateField = '收款日期';
const findCombination = (target, availableInvoices) => {
  if (target > availableInvoices.reduce((amount, inv) => amount + inv[billAmountField], 0)) {
    return null;
  }
  // 回溯法尝试找一组发票金额组合，正好等于 target
  const result = [];
  let found = false;
  const dfs = (start, path, sum) => {
    if (found) return;
    if (sum === target) {
      result.push(...path);
      found = true;
      return;
    }
    if (sum > target) return;
    for (let i = start; i < availableInvoices.length; i++) {
      const inv = availableInvoices[i];
      if (inv.used) continue;
      path.push(inv);
      dfs(i + 1, path, sum + inv[billAmountField]);
      path.pop();
    }
  };
  dfs(0, [], 0);
  return found ? result : null;
};

ipcMain.handle('startMerge', async (_, offset) => {
  try {
    const billData = utils.xlsxParser(billFilePath);
    const billDataHeaderRow = billData[0];
    if (!billDataHeaderRow?.includes(billCodeField)) {
      return '发票数据表中没有发票号码数据';
    }
    if (!billDataHeaderRow?.includes(billAmountField)) {
      return '发票数据表中没有发票金额数据';
    }
    if (!billDataHeaderRow?.includes(billDateField)) {
      return '发票数据表中没有发票日期数据';
    }
    const receiptData = utils.xlsxParser(receiptFilePath);
    const receiptDataHeaderRow = receiptData[0];
    if (!receiptDataHeaderRow?.includes(receiptAmountField)) {
      return '收款数据表中没有收款金额数据';
    }
    if (!receiptDataHeaderRow?.includes(receiptDateField)) {
      return '收款数据表中没有收款日期数据';
    }
    const operator = profile.licenseInfo.username;
    const currentTime = new Date().toLocaleString();
    const resolvedData1 = [];
    const resolvedData2 = [];
    const resolvedBillData = utils.dataResolver(billData).filter(item => item[billAmountField]);
    utils.dataResolver(receiptData).forEach(receipt => {
      const receiptAmount = receipt[receiptAmountField];
      const receiptDate = new Date(receipt[receiptDateField]);
      const availableInvoices = resolvedBillData.filter(item => !item.used && (receiptDate - new Date(item[billDateField]) >= 86400000 * offset)); // 过滤：未使用 & 发票日期 < 收款日期
      const match = availableInvoices.length ? findCombination(receiptAmount, availableInvoices) : null;
      if (match) {
        match.forEach(m => (m.used = true));
        resolvedData1.push([receiptAmount, receipt[receiptDateField], match.map(item => item[billCodeField]).join(', '), operator, currentTime]);
        match.forEach(item => resolvedData2.push([item[billCodeField], item[billAmountField], item[billDateField], receiptAmount, receipt[receiptDateField], operator, currentTime]));
      } else {
        resolvedData1.push([receiptAmount, receipt[receiptDateField], '没有找到与当前收款金额对应的发票组合', operator, currentTime]);
      }
    });
    resolvedBillData.filter(item => !item.used).forEach(item => resolvedData2.push([item[billCodeField], item[billAmountField], item[billDateField]]));
    resolvedData1.unshift([...receiptDataHeaderRow, '发票代码组', operatorField, '操作时间']);
    resolvedData2.unshift([billCodeField, billAmountField, billDateField, receiptAmountField, receiptDateField, operatorField, '操作时间']);
    await utils.xlsxSaver(mergeResultFilePath1, [{
      name: '聚合结果表',
      data: resolvedData1
    }], [{ wch: 20 }, { wch: 20 }, { wch: 60 }, { wch: 20 }, { wch: 20 }]);
    await utils.xlsxSaver(mergeResultFilePath2, [{
      name: '聚合结果表',
      data: resolvedData2
    }], [{ wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }]);
    return '聚合发票完成';
  } catch (error) {
    return `聚合发票数据失败：${ error.message }`;
  }
});

ipcMain.handle('openFile', (_, filePath) => shell.openPath(`${ dataPath }/${ filePath }`));
