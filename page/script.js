const validateLogin = async () => {
    const { loginTime } = await $contextBridge.getLoginInfo();
    if (Date.now() - loginTime > 2 * 3600 * 1000) {
        $utils.setToast('请重新登录');
        logoutLink.click();
        return false;
    }
    return true;
};

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
        const isLoginValid = await validateLogin();
        isLoginValid && importPdfFileInput.click();
    });
    importPdfFileInput.addEventListener('change', () => {
        importPdfFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入pdf文件并识别文字内容...', 180000);
        $contextBridge.importPdfFile(importPdfFileInput.files[0], radio1.checked).then(message => {
            importPdfFileButton.removeAttribute('disabled', '');
            importPdfFileInput.value = '';
            $utils.setToast(message);
        });
    });
    openPdfFileButton.addEventListener('click', () => $contextBridge.openFile('input.pdf'));
    openExcelFileButton.addEventListener('click', () => $contextBridge.openFile('output.xlsx'));
    // 订单拆分功能
    document.querySelector('#typeSelect').addEventListener('change', ({ target }) => {
        panel12.setAttribute('style', 'display: none');
        panel3.setAttribute('style', 'display: none');
        panel4.setAttribute('style', 'display: none');
        if ([radio1, radio2].includes(target)) {
            panel12.setAttribute('style', 'display: block');
        } else if (target === radio3) {
            panel3.setAttribute('style', 'display: block');
        } else if (target === radio4) {
            panel4.setAttribute('style', 'display: block');
        }
    });
    // 导入产品类别表按钮点击事件
    importProductTypeFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importProductTypeFileInput.click();
    });
    importProductTypeFileInput.addEventListener('change', () => {
        importProductTypeFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入产品类别表...', 180000);
        $contextBridge.importProductTypeFile(importProductTypeFileInput.files[0]).then(message => {
            importProductTypeFileButton.removeAttribute('disabled', '');
            importProductTypeFileInput.value = '';
            $utils.setToast(message);
        });
    });
    // 导入产品价格表按钮点击事件
    importPriceFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importPriceFileInput.click();
    });
    importPriceFileInput.addEventListener('change', () => {
        importPriceFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入产品价格表...', 180000);
        $contextBridge.importPriceFile(importPriceFileInput.files[0]).then(message => {
            importPriceFileButton.removeAttribute('disabled', '');
            importPriceFileInput.value = '';
            $utils.setToast(message);
        });
    });
    // 导入经销商信息表按钮点击事件
    importDealerFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importDealerFileInput.click();
    });
    importDealerFileInput.addEventListener('change', () => {
        importDealerFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入经销商信息表...', 180000);
        $contextBridge.importDealerFile(importDealerFileInput.files[0]).then(message => {
            importDealerFileButton.removeAttribute('disabled', '');
            importDealerFileInput.value = '';
            $utils.setToast(message);
        });
    });
    // 导入原始订单数据表按钮点击事件
    importOrderFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importOrderFileInput.click();
    });
    importOrderFileInput.addEventListener('change', () => {
        importOrderFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入原始订单数据表并拆分订单...', 180000);
        $contextBridge.importOrderFile(importOrderFileInput.files[0]).then(message => {
            importOrderFileButton.removeAttribute('disabled', '');
            importOrderFileInput.value = '';
            $utils.setToast(message);
        });
    });
    openProductTypeFileButton.addEventListener('click', () => $contextBridge.openFile('productTypes.xlsx'));
    openPriceFileButton.addEventListener('click', () => $contextBridge.openFile('prices.xlsx'));
    openDealerFileButton.addEventListener('click', () => $contextBridge.openFile('dealers.xlsx'));
    openOrderFileButton.addEventListener('click', () => $contextBridge.openFile('orders.xlsx'));
    openSplitOrderFileButton.addEventListener('click', () => $contextBridge.openFile('splitOrders.xlsx'));
    openInvalidOrderFileButton.addEventListener('click', () => $contextBridge.openFile('invalidOrders.xlsx'));
    // 发票聚合功能
    // 导入发票列表按钮点击事件
    importBillFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importBillFileInput.click();
    });
    importBillFileInput.addEventListener('change', () => {
        importBillFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入发票列表文件...', 180000);
        $contextBridge.importBillFile(importBillFileInput.files[0]).then(message => {
            importBillFileButton.removeAttribute('disabled', '');
            importBillFileInput.value = '';
            $utils.setToast(message);
        });
    });
    // 导入收款列表按钮点击事件
    importReceiptFileButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && importReceiptFileInput.click();
    });
    importReceiptFileInput.addEventListener('change', () => {
        importReceiptFileButton.setAttribute('disabled', '');
        $utils.setToast('正在导入收款列表文件...', 180000);
        $contextBridge.importReceiptFile(importReceiptFileInput.files[0]).then(message => {
            importReceiptFileButton.removeAttribute('disabled', '');
            importReceiptFileInput.value = '';
            $utils.setToast(message);
        });
    });
    startMergeButton.addEventListener('click', async () => {
        const isLoginValid = await validateLogin();
        isLoginValid && $contextBridge.startMerge().then(message => $utils.setToast(message)); 
    })
    openBillFileButton.addEventListener('click', () => $contextBridge.openFile('bill.xlsx'));
    openReceiptFileButton.addEventListener('click', () => $contextBridge.openFile('receipt.xlsx'));
    openMergeResultFileButton1.addEventListener('click', () => $contextBridge.openFile('mergeResult1.xlsx'));
    openMergeResultFileButton2.addEventListener('click', () => $contextBridge.openFile('mergeResult2.xlsx'));
    $utils.toggleView('loginView');
});
