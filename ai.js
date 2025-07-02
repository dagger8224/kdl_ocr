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