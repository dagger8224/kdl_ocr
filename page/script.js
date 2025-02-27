
window.addEventListener("DOMContentLoaded", async () => {
    // 获取配置信息
    const version = await $contextBridge.getVersion();
    if (version > $contextBridge.version) {
        $utils.setToast('当前版本不再支持，请与客服联系！');
        errorView.innerHTML = `当前版本不再支持，请发邮件到<a href="mailto:174876467@qq.com?subject=【DeepSeek订单智能处理系统】版本更新">174876467@qq.com</a>获取应用程序最新版本。`;
        return $utils.toggleView('errorView');
    }

    // 登录按钮点击事件
    loginButton.addEventListener('click', async event => {
        event.preventDefault();
        const username = userNameInput.value.trim();
        if (!username) {
            return $utils.setToast('请输入用户名');
        }
        const password = passwordInput.value.trim();
        if (!password) {
            return $utils.setToast('请输入密码');
        }
        const { message, licenseInfo } = await $contextBridge.verifyLicense(username, password);
        $utils.setToast(message);
        if (!licenseInfo) { return; }
        const { loginTime, leftDays } = licenseInfo;
        if (leftDays) {
            $utils.setLoginDuration(loginTime);
            return $utils.toggleView('operateView');
        } else {
            $utils.setToast('当前账号未获得有效授权，请与客服联系！');
            errorView.innerHTML = `用户“${ username }”当前尚未获得授权或者授权条件已终止，请发邮件到<a href="mailto:174876467@qq.com?subject=【阳采平台数据智能处理程序】授权问题咨询">174876467@qq.com</a>沟通授权相关问题。`;
            $utils.toggleView('errorView');
            return false;
        }
    });

    // 退出登录按钮点击事件
    logoutLink.addEventListener('click', async () => {
        await $contextBridge.logout();
        $utils.toggleView('loginView');
    });

    // 导入pdf文件按钮点击事件
    importPdfFileButton.addEventListener('click', async () => {
        const { loginTime } = await $contextBridge.getLoginInfo();
        if (Date.now() - loginTime > 2 * 3600 * 1000) {
            $utils.setToast('请重新登录');
            return logoutLink.click();
        }
        importPdfFileInput.click();
    });
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
    $utils.toggleView('loginView');
});
