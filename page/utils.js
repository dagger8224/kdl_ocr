const viewIdList = ['errorView', 'loginView', 'operateView'];

let $timerId = 0;

window.$utils = {
    setLoginDuration: startTime => {
        $timerId && clearTimeout($timerId);
        $timerId = setTimeout(() => $utils.setLoginDuration(startTime), 1000);
        const seconds = (Date.now() - startTime) / 1000;
        const hours = Math.floor(seconds / 3600);
        const minutes = Math.floor((seconds % 3600) / 60);
        const remainingSeconds = Math.floor(seconds % 60);
        const formattedHours = hours.toString().padStart(2, '0');
        const formattedMinutes = minutes.toString().padStart(2, '0');
        const formattedSeconds = remainingSeconds.toString().padStart(2, '0');
        loginDuration.innerText = `${formattedHours}:${formattedMinutes}:${formattedSeconds}`;
    },
    setToast: (message, duration = 2000) => {
        toast.innerText = message;
        toast.setAttribute('visible', '');
        setTimeout(() => {
            toast.innerText = '';
            toast.removeAttribute('visible');
        }, duration);
        return false;
    },
    setVerifyCode: async () => {
        const { codeImage, message } = await $contextBridge.getVerifyCode();
        if (codeImage) {
            verifyCode.setAttribute('src', codeImage);
        } else {
            $utils.setToast(message);
        }
        return codeImage;
    },
    toggleView: (viewId, visible = true) => {
        viewIdList.forEach(viewId => window[viewId].removeAttribute('visible'));
        visible && window[viewId].setAttribute('visible', '');
    }
};
