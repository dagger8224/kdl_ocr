// entry

const fs = require('fs');
const path = require('path');
const { app, BrowserWindow, ipcMain, shell } = require('electron');
const utils = require('./utils.js');

const dataPath = '../data';
const inputFilePath = `${ dataPath }/input.pdf`;
const outputFilePath = `${ dataPath }/output.xlsx`;

const loginErrorMessage = '账号或者密码有误，请重新输入';

const profile = {
  config: {},
  licenseInfo: {
    loginTime: 0
  }
};

// 观察文件状态
const fileExistsStateWatcher = () => BrowserWindow.getAllWindows()[0].webContents.send('fileExistsState', {
  openPdfFileButton: fs.existsSync(inputFilePath),
  openExcelFileButton: fs.existsSync(outputFilePath),
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
    const headerRow = isType1 ? ['发货日期', '发货单号', '客户订单号', '西门子订单号', '西门子编码', '储存温度'] : ['发货日期', '客户订单号', '西门子订单号', '西门子编码', '储存温度', '数量', '单价', '金额'];
    const resultData = [headerRow];
    const groupPrefix = ['批号', '数量', '单位', '效期', '生产日期']; // 可能有多组批号效期数据
    const groupLength = groupPrefix.length;
    let baseIndex = isType1 ? 0 : 11;
    let itemStartNumber = fullTokens[baseIndex] - 0; // 起始序号转为数字，应该为10/1000
    let itemStartNumberIndex = 0;
    while (fullTokens.length) {
      const rowData = [...baseRowData];
      rowData.push(fullTokens[baseIndex + 1]); // 西门子编码
      itemStartNumber += isType1 ? 10 : 1000;
      itemStartNumberIndex = fullTokens.indexOf((itemStartNumber + '').padStart(6, '0'));
      itemStartNumberIndex === -1 && (itemStartNumberIndex = fullTokens.length);
      const currentTokens = fullTokens.slice(0, itemStartNumberIndex);
      const storageCondition = currentTokens.find(token => token.trim().startsWith('Storage Condition:'));
      rowData.push(storageCondition ? storageCondition.replace('Storage Condition:', '').trim() : '-');
      if (isType1) {
        // 第一页结构与后面的不一样，分别处理
        startIndex = 0;
        if (itemStartNumber === 20) {
          startIndex = currentTokens.findIndex(token => token.trim().endsWith('Manuf. Date'));
        } else {
          startIndex = currentTokens.findIndex(token => token.trim().startsWith('CFDA License:'));
          if (startIndex === -1) {
            startIndex = currentTokens.findIndex(token => token.trim().startsWith('Country of Origin:'));
          }
        }
        const splitTokens = currentTokens.slice(startIndex + 1, itemStartNumberIndex).join('/').split('/').filter(token => token.trim()).map(token => token.trim());
        splitTokens.forEach((token, index) => {
          const fieldName = `${ groupPrefix[index % groupLength] }${ Math.floor(index / groupLength) + 1 }`;
          headerRow.includes(fieldName) || headerRow.push(fieldName);
          if ((index === 3 || index === 4) && !token.includes('/')) {
            token = `${ token.slice(0, 4) }/${ token.slice(4, 6) }/${ token.slice(6, 8) }`;
          }
          rowData.push(token);
        });
      } else {
        rowData.push(fullTokens[baseIndex + 3]); // 数量
        rowData.push(fullTokens[baseIndex + 4]); // 单价
        rowData.push(fullTokens[baseIndex + 5]); // 金额
      }
      resultData.push(rowData);
      fullTokens = fullTokens.slice(itemStartNumberIndex);
      baseIndex = 0;
    }
    utils.xlsxSaver(outputFilePath, '识别结果', resultData, headerRow.map(() => ({ wch: 16 })));
    return '内容识别完成';
  } catch (error) {
    return error.code === 'EBUSY' ? `请关闭已打开的pdf和转换结果文件后再试` : `操作失败：${ error.message }`;
  }
});

ipcMain.handle('openFile', (_, filePath) => shell.openPath(`${ dataPath }/${ filePath }`));
