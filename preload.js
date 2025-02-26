// 向渲染进程注入数据
const { contextBridge, ipcRenderer, webUtils } = require('electron');

window.addEventListener('DOMContentLoaded', () => {
  const version = '1.00';
  document.title = '康德乐订单智能处理系统V1.0.0';
  contextBridge.exposeInMainWorld('$contextBridge', {
    version,
    getVersion: () => ipcRenderer.invoke('getVersion'),
    importPdfFile: file => file && ipcRenderer.invoke('importPdfFile', webUtils.getPathForFile(file)),
    openFile: filePath => ipcRenderer.invoke('openFile', filePath),
    startConvert: () => ipcRenderer.invoke('startConvert')
  });
  ipcRenderer.on('fileExistsState', (_, states) => {
    ['openPdfFileButton', 'openExcelFileButton'].forEach(id => states[id] ? window[id].removeAttribute('disabled') : window[id].setAttribute('disabled', ''));
  });
});
