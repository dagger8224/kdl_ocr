
window.addEventListener("DOMContentLoaded", async () => {
    // 获取配置信息
    /* const version = await $contextBridge.getVersion();
    if (version > $contextBridge.version) {
        $utils.setToast('当前版本不再支持，请与客服联系！');
        errorView.innerHTML = `当前版本不再支持，请发邮件到<a href="mailto:174876467@qq.com?subject=【康德乐订单智能处理系统】版本更新">174876467@qq.com</a>获取应用程序最新版本。`;
        return $utils.toggleView('errorView');
    } */
    // 导入pdf文件按钮点击事件
    importPdfFileButton.addEventListener('click', () => importPdfFileInput.click());
    importPdfFileInput.addEventListener('change', () => {
        importPdfFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入pdf文件并识别文字内容...', 180000);
        $contextBridge.importPdfFile(importPdfFileInput.files[0]).then(message => {
            importPdfFileButton.removeAttribute('disabled', '');
            importPdfFileInput.value = '';
            $utils.setToast(message);
        });
    });
    openPdfFileButton.addEventListener('click', () => $contextBridge.openFile('input.pdf'));
    openExcelFileButton.addEventListener('click', () => $contextBridge.openFile('output.xlsx'));
});
