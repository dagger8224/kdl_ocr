// 向渲染进程注入数据
const { contextBridge, ipcRenderer, webUtils } = require('electron');

window.addEventListener('DOMContentLoaded', () => {
  const version = '1.00';
  document.title = 'DeepSeek订单智能处理系统V1.0.0';
  contextBridge.exposeInMainWorld('$contextBridge', {
    version,
    getVersion: () => ipcRenderer.invoke('getVersion'),
    getLoginInfo: () => ipcRenderer.invoke('getLoginInfo'),
    verifyLicense: (account, password) => ipcRenderer.invoke('verifyLicense', account, password),
    login: form => ipcRenderer.invoke('login', form),
    logout: () => ipcRenderer.invoke('logout'),
    importPdfFile: file => file && ipcRenderer.invoke('importPdfFile', webUtils.getPathForFile(file)),
    openFile: filePath => ipcRenderer.invoke('openFile', filePath),
    startConvert: () => ipcRenderer.invoke('startConvert')
  });
  ipcRenderer.on('fileExistsState', (_, states) => {
    ['openPdfFileButton', 'openExcelFileButton'].forEach(id => states[id] ? window[id].removeAttribute('disabled') : window[id].setAttribute('disabled', ''));
  });
  ipcRenderer.on('updateHeader', (_, profileInfo) => {
    const { username, leftDays, authMode } =  profileInfo;
    window.headerUserName.innerText = `${ username } [软件版本 V1.0.0 剩余授权时间：${ leftDays }天${ authMode }]`;
  });
});
