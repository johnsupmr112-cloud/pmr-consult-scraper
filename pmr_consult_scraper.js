/* =========================================================================
 * PMR 會診清單擷取腳本  v4  →  直接下載排版好的 Excel
 *
 * 輸出：PMR_YYYYMMDD.xlsx
 *   欄位 病房床號 / 病歷號 / 病人姓名 / 會診內容 / 病人科別 / 會診時限 / 被會醫師
 *   新細明體 12pt、全部置中、會診內容自動換行、列高自動調整
 *   A4 橫式、縮放至一頁寬、表頭每頁重複 → 開啟後直接 Ctrl+P 列印
 *
 * 用法：
 *   1. 在會診系統查詢出清單頁面
 *   2. F12 → Console → 貼上整段 → Enter
 *   3. 等進度跑完，瀏覽器自動下載 xlsx
 *
 * ★ 第一次使用請先設 CONFIG.diagnose = true 確認「病史」欄位抓對位置。
 * ========================================================================= */

(async () => {
  'use strict';

  // ======================= 可調整參數 =======================
  const CONFIG = {
    columns: [
      // [輸出標題, 清單原始欄位（null=從明細頁抓）, 欄寬, 是否自動換行]
      ['病房床號', '病房床號', 10.5, false],
      ['病歷號', '病歷號', 10, false],
      ['病人姓名', '病人姓名', 10, false],
      ['會診內容', null, 95, true],
      ['病人科別', '病人科別', 9.5, false],
      ['會診時限', '會診時限', 15.5, false],
      ['被會醫師', '被會醫師', 13, false],
    ],

    fontName: '新細明體',
    fontSize: 12,
    contentAlign: 'left',   // 會診內容「內文」的水平對齊（標題永遠置中）
    contentVAlign: 'top',   // 會診內容的垂直對齊。多行段落用 'top' 較好讀，也可填 'center'
    preserveLineBreaks: true, // true = 保留段落結構；false = 全部壓成一段
    unwrapHardBreaks: true,   // true = 還原「固定寬度硬切」造成的詭異斷行
    joinBrokenParagraphs: true, // true = 合併把句子切成兩半的假空行
    boldHeader: true,       // 表頭粗體
    headerFill: 'E2EFDA',   // 表頭底色。想換成灰色填 'D9D9D9'，不要底色填 null
    borderColor: '000000',  // 框線顏色。想淡一點填 'BFBFBF'

    // 列印相關（若列印仍有問題，可逐一改成 false 測試）
    freezeHeader: true,           // 凍結首列
    repeatHeaderOnEveryPage: true, // 每頁重複表頭

    fetchDetail: true,
    detailLabels: ['病史', '會診原因', '會診目的'],
    delayMs: 400,
    maxRows: 0,             // 0 = 不限；測試時可設 3
    diagnose: false,        // ★ 只抓第一筆並印出明細頁結構

    onlyPMR: true,
    statusFilter: null,     // null / '待會' / 'V'
    sortBy: '會診時限',
    sortDesc: true,
    dedupe: true,
  };
  // =========================================================

  // 單行用：壓成一行（清單欄位如床號、姓名）
  const clean = (s) =>
    (s || '')
      .replace(/\u00a0/g, ' ')
      .replace(/[\t\r\n]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const txt = (el) => clean(el ? el.textContent : '');

  /**
   * 還原固定寬度硬換行。
   * 有些系統把病史以每 N 字元硬切一行的方式儲存，會切在單字中間
   * （例：sensat / ion）。這裡偵測出切斷寬度後把它們接回原句。
   */
  const unwrapHard = (text) => {
    const lines = text.split('\n');
    if (lines.length < 3) return text;

    // 用「出現最多次的行長」判定切斷寬度，比取最大值穩健
    // （偶爾有一行特別長就會讓最大值失準，導致該接的地方被插入多餘空白）
    const freq = new Map();
    for (const l of lines) if (l.length >= 40) freq.set(l.length, (freq.get(l.length) || 0) + 1);
    if (!freq.size) return text;
    let W = 0;
    let best = 0;
    for (const [len, n] of freq) {
      if (n > best || (n === best && len > W)) {
        best = n;
        W = len;
      }
    }
    if (W < 40 || best < 2) return text;

    const out = [];
    let buf = null; // { text, prevLen }
    const flush = () => {
      if (buf !== null) {
        out.push(buf.text);
        buf = null;
      }
    };

    for (const line of lines) {
      if (!line.trim()) {
        flush();
        out.push('');
        continue;
      }
      if (buf === null) {
        buf = { text: line, prevLen: line.length };
      } else {
        // 前一行剛好被切滿 → 原字元已含空白（或切在字中間），直接相接
        // 前一行短一點 → 行尾空白已被去掉，補回一個空白
        const needSpace =
          buf.prevLen < W && !/\s$/.test(buf.text) && !/^\s/.test(line);
        buf.text += (needSpace ? ' ' : '') + line;
        buf.prevLen = line.length;
      }
      // 這一行沒被切滿 → 段落到此結束
      if (line.length < W - 2) flush();
    }
    flush();

    return out.join('\n');
  };

  /**
   * 合併「假的空行」。
   * 硬換行的資料常在句子中間夾一個空行，把一句話切成兩段。
   * 只有在前段結尾沒有句尾標點、且後段以小寫字母開頭時才合併，
   * 避免誤把真正的段落或條列接在一起。
   */
  const joinBroken = (lines) => {
    const out = [];
    for (let i = 0; i < lines.length; i++) {
      const cur = lines[i];
      if (
        cur &&
        lines[i + 1] === '' &&
        lines[i + 2] &&
        !/[.!?:;,]$/.test(cur) &&        // 前段不是句尾
        /^[a-z(]/.test(lines[i + 2])     // 後段以小寫或左括號開頭
      ) {
        lines[i + 2] = cur + ' ' + lines[i + 2];
        i++; // 跳過空行，cur 已併入後段
        continue;
      }
      out.push(cur);
    }
    return out;
  };

  // 多行用：保留真正的段落結構（會診內容）
  const cleanML = (s) => {
    let t = (s || '').replace(/\u00a0/g, ' ').replace(/\r\n?/g, '\n');
    if (CONFIG.unwrapHardBreaks) t = unwrapHard(t);

    let lines = t
      .split('\n')
      .map((line) => line.replace(/[ \t]+/g, ' ').trim());

    if (CONFIG.joinBrokenParagraphs) lines = joinBroken(lines);

    return lines
      .join('\n')
      .replace(/\n{3,}/g, '\n\n') // 連續空行最多留一行
      .replace(/^\n+|\n+$/g, ''); // 去掉頭尾空行
  };

  /**
   * 把元素轉成保留段落結構的文字。
   * <br> → 換行；<p>/<div>/<tr>/<li> 等區塊元素前後 → 換行。
   */
  const blockText = (el) => {
    if (!el) return '';
    // textarea / input 的值本身就含換行，直接取
    if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT')
      return cleanML(el.value || el.textContent || el.getAttribute('value') || '');

    const BLOCK = /^(P|DIV|TR|LI|BR|H[1-6]|BLOCKQUOTE|PRE|TABLE|UL|OL|SECTION|ARTICLE)$/;
    let out = '';
    const walk = (node) => {
      for (const n of node.childNodes) {
        if (n.nodeType === 3) {
          out += n.nodeValue;
        } else if (n.nodeType === 1) {
          const tag = n.tagName;
          if (tag === 'BR') {
            out += '\n';
            continue;
          }
          if (tag === 'SCRIPT' || tag === 'STYLE') continue;
          // <pre> 內的空白原樣保留
          if (tag === 'PRE') {
            out += '\n' + n.textContent + '\n';
            continue;
          }
          // 區塊元素前後各補一個換行，但不重複疊加
          // （<br><br> 造成的空行仍會保留）
          const nl = () => {
            if (!/\n$/.test(out)) out += '\n';
          };
          if (BLOCK.test(tag)) nl();
          walk(n);
          if (BLOCK.test(tag)) nl();
        }
      }
    };
    walk(el);
    return cleanML(out);
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* ===================== XLSX 產生器（純 JS，無外部相依） ===================== */
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };

  function zipStore(files) {
    const enc = new TextEncoder();
    const parts = [];
    const central = [];
    let offset = 0;
    const u16 = (n) => [n & 0xff, (n >>> 8) & 0xff];
    const u32 = (n) => [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff];
    const modTime = 0;
    const modDate = ((2020 - 1980) << 9) | (1 << 5) | 1;

    for (const f of files) {
      const nameBytes = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);
      const local = [
        ...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(0),
        ...u16(modTime), ...u16(modDate),
        ...u32(crc), ...u32(data.length), ...u32(data.length),
        ...u16(nameBytes.length), ...u16(0),
      ];
      parts.push(new Uint8Array(local), nameBytes, data);
      const cd = [
        ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(0),
        ...u16(modTime), ...u16(modDate),
        ...u32(crc), ...u32(data.length), ...u32(data.length),
        ...u16(nameBytes.length), ...u16(0), ...u16(0),
        ...u16(0), ...u16(0), ...u32(0), ...u32(offset),
      ];
      cd.nameBytes = nameBytes;
      central.push(cd);
      offset += local.length + nameBytes.length + data.length;
    }

    const cdParts = [];
    let cdSize = 0;
    for (const c of central) {
      const head = new Uint8Array(c);
      cdParts.push(head, c.nameBytes);
      cdSize += head.length + c.nameBytes.length;
    }
    const eocd = new Uint8Array([
      ...u32(0x06054b50), ...u16(0), ...u16(0),
      ...u16(files.length), ...u16(files.length),
      ...u32(cdSize), ...u32(offset), ...u16(0),
    ]);

    const all = [...parts, ...cdParts, eocd];
    const out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
    let p = 0;
    for (const a of all) {
      out.set(a, p);
      p += a.length;
    }
    return out;
  }

  const esc = (s) =>
    String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      // 換行寫成實體，避免 XML 剖析時被正規化成空白
      .replace(/\r?\n/g, '&#10;');

  const colLetter = (n) => {
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  function buildXlsx(cols, rows, sheetName) {
    const headers = cols.map((c) => c[0]);
    const font = CONFIG.fontName;
    const size = CONFIG.fontSize;
    sheetName = String(sheetName).replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 31);

    const bc = CONFIG.borderColor || '000000';
    const side = `<color rgb="FF${bc}"/>`;
    const hdrFill = CONFIG.headerFill
      ? `<patternFill patternType="solid"><fgColor rgb="FF${CONFIG.headerFill}"/><bgColor indexed="64"/></patternFill>`
      : `<patternFill patternType="none"/>`;

    // fillId 2 = 表頭底色；borderId 1 = 四邊細框線
    // cellXfs：1=一般置中+框線  2=內文換行+框線  3=表頭
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="${size}"/><name val="${esc(font)}"/><family val="2"/><charset val="136"/></font>
<font><b/><sz val="${size}"/><name val="${esc(font)}"/><family val="2"/><charset val="136"/></font>
</fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill>${hdrFill}</fill></fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin">${side}</left><right style="thin">${side}</right><top style="thin">${side}</top><bottom style="thin">${side}</bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="${CONFIG.contentAlign}" vertical="${CONFIG.contentVAlign}" wrapText="1"/></xf>
<xf numFmtId="0" fontId="${CONFIG.boldHeader ? 1 : 0}" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

    const colsXml = cols
      .map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${c[2]}" customWidth="1"/>`)
      .join('');

    const cell = (r, ci, val, s) =>
      `<c r="${colLetter(ci)}${r}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(
        val
      )}</t></is></c>`;

    // ---- 預先算好列高 ----
    // 讓 Excel 在列印時不必臨場重算整份版面的自動列高
    const dispWidth = (s) => {
      let w = 0;
      for (const ch of String(s)) w += ch.codePointAt(0) >= 0x2e80 ? 2 : 1; // 中日韓字元算 2
      return w;
    };
    const LINE_PT = CONFIG.fontSize * 1.375; // 12pt → 16.5pt，與 Excel 預設一致
    const rowHeight = (row) => {
      let maxLines = 1;
      for (const c of cols) {
        if (!c[3]) continue; // 只有自動換行的欄位才需要估算
        const avail = Math.max(4, c[2] - 1); // 扣掉儲存格左右內距
        let lines = 0;
        for (const seg of String(row[c[0]] ?? '').split('\n'))
          lines += Math.max(1, Math.ceil(dispWidth(seg) / avail));
        maxLines = Math.max(maxLines, lines);
      }
      // Excel 單列高度上限為 409.5pt
      return Math.min(409, Math.round(maxLines * LINE_PT * 10) / 10);
    };

    let sd = `<row r="1" ht="${LINE_PT * 1.5}">${headers
      .map((h, i) => cell(1, i + 1, h, 3))
      .join('')}</row>`;
    rows.forEach((row, ri) => {
      const r = ri + 2;
      // 給定 ht 但不加 customHeight，Excel 仍可在編輯後自行調整
      sd += `<row r="${r}" ht="${rowHeight(row)}">${cols
        .map((c, i) => cell(r, i + 1, row[c[0]] ?? '', c[3] ? 2 : 1))
        .join('')}</row>`;
    });

    const qName = `'${esc(sheetName).replace(/'/g, "''")}'`;
    const titlesXml = CONFIG.repeatHeaderOnEveryPage
      ? `<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">${qName}!$1:$1</definedName></definedNames>`
      : '';
    const paneXml = CONFIG.freezeHeader
      ? `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>`
      : `<selection activeCell="A2" sqref="A2"/>`;

    const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:${colLetter(headers.length)}${rows.length + 1}"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0">${paneXml}</sheetView></sheetViews>
<sheetFormatPr baseColWidth="8" defaultRowHeight="16.5"/>
<cols>${colsXml}</cols>
<sheetData>${sd}</sheetData>
<printOptions horizontalCentered="0" verticalCentered="0"/>
<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>
<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>
</worksheet>`;

    const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<workbookPr/>
<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="12000" firstSheet="0" activeTab="0"/></bookViews>
<sheets><sheet name="${esc(sheetName)}" sheetId="1" state="visible" r:id="rId1"/></sheets>
${titlesXml}
<calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>`;

    // Excel 的樣式引擎會去解析佈景主題，缺少這個部件會讓完整版面計算（例如列印）失敗
    const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

    const iso = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

    return zipStore([
      {
        name: '[Content_Types].xml',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,
      },
      {
        name: '_rels/.rels',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,
      },
      {
        name: 'docProps/core.xml',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>${esc(sheetName)}</dc:title>
<dcterms:created xsi:type="dcterms:W3CDTF">${iso}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified>
</cp:coreProperties>`,
      },
      {
        name: 'docProps/app.xml',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Microsoft Excel</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>`,
      },
      { name: 'xl/workbook.xml', data: workbook },
      {
        name: 'xl/_rels/workbook.xml.rels',
        data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`,
      },
      { name: 'xl/theme/theme1.xml', data: theme },
      { name: 'xl/styles.xml', data: styles },
      { name: 'xl/worksheets/sheet1.xml', data: sheet },
    ]);
  }
  /* =================== XLSX 產生器結束 =================== */

  // ---------- 1. 找出資料表 ----------
  const tables = [...document.querySelectorAll('table')].filter(
    (t) => t.querySelectorAll('table').length === 0
  );
  if (!tables.length) {
    console.error('[PMR] 找不到 table。表格可能在 iframe 內，請先切換 frame。');
    return;
  }

  let table = tables.find((t) =>
    [...t.rows].some((tr) => [...tr.cells].some((c) => clean(c.innerText).includes('病歷號')))
  );
  if (!table) {
    table = tables
      .map((t) => ({ t, s: t.rows.length * Math.max(1, ...[...t.rows].map((r) => r.cells.length)) }))
      .sort((a, b) => b.s - a.s)[0].t;
    console.warn('[PMR] 找不到含「病歷號」的表頭，改用最大表格。');
  }

  const trs = [...table.rows];
  const grid = trs.map((tr) => [...tr.cells].map((td) => clean(td.innerText)));
  const hIdx = grid.findIndex((r) => r.some((c) => c.includes('病歷號')));
  if (hIdx === -1) {
    console.error('[PMR] 找不到標題列。');
    return;
  }
  const rawHeaders = grid[hIdx];
  console.log('[PMR] 系統原始表頭：', rawHeaders);

  const findCol = (want) => {
    if (!want) return -1;
    let i = rawHeaders.findIndex((h) => h === want);
    if (i === -1) i = rawHeaders.findIndex((h) => h && h.includes(want));
    return i;
  };
  const map = CONFIG.columns.map(([out, src]) => ({ out, idx: findCol(src) }));
  const iDept = findCol('被會科別');
  const iStat = findCol('狀態');

  // ---------- 2. 過濾 ----------
  let items = [];
  for (let i = hIdx + 1; i < grid.length; i++) {
    const cells = grid[i];
    if (!cells.join('').trim()) continue;
    if (CONFIG.onlyPMR && iDept !== -1 && !(cells[iDept] || '').toUpperCase().includes('PMR'))
      continue;
    if (CONFIG.statusFilter && iStat !== -1) {
      const s = (cells[iStat] || '').toUpperCase();
      const want = CONFIG.statusFilter.toUpperCase();
      if (want === 'V' ? s !== 'V' : !s.includes(CONFIG.statusFilter)) continue;
    }
    const obj = {};
    for (const { out, idx } of map) obj[out] = idx === -1 ? '' : cells[idx] || '';
    if (!Object.values(obj).join('').trim()) continue;
    items.push({ obj, tr: trs[i] });
  }

  if (CONFIG.dedupe) {
    const seen = new Set();
    items = items.filter(({ obj }) => {
      const k = `${obj['病歷號']}|${obj['會診時限']}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }
  if (CONFIG.maxRows > 0) items = items.slice(0, CONFIG.maxRows);

  console.log(`[PMR] 待處理 ${items.length} 筆`);
  if (!items.length) {
    console.warn('[PMR] 沒有符合條件的資料。試著把 onlyPMR 改成 false。');
    return;
  }

  // ---------- 3. 明細頁連結 ----------
  const linkOf = (tr) => {
    const anchors = [...tr.querySelectorAll('a')];
    const a = anchors.find((x) => (x.textContent || '').includes('查詢')) || anchors[0];
    if (!a) return null;
    const href = a.getAttribute('href') || '';
    if (href && !/^\s*(javascript:|#)/i.test(href)) return new URL(href, location.href).href;
    const src = `${a.getAttribute('onclick') || ''} ${href}`;
    const m =
      src.match(/['"]([^'"]*\.cfm[^'"]*)['"]/i) ||
      src.match(/['"]([^'"]*\?[^'"]*(?:ordno|seqno|consno|id)=[^'"]*)['"]/i);
    return m ? new URL(m[1], location.href).href : null;
  };

  // ---------- 4. 明細欄位擷取 ----------
  const extractField = (doc, labels) => {
    const cells = [...doc.querySelectorAll('td, th')];
    for (const label of labels) {
      // (a) 標題儲存格 → 取右邊（或下一列同位置）的內容，保留換行
      for (const c of cells) {
        if (txt(c).replace(/[:：\s]/g, '') !== label) continue;
        let n = c.nextElementSibling;
        while (n && !txt(n)) n = n.nextElementSibling;
        if (n) {
          const ta = n.querySelector('textarea, input[type=text], input:not([type])');
          const v = blockText(ta || n);
          if (v) return v;
        }
        const tr = c.closest('tr');
        const idx = [...(tr ? tr.cells : [])].indexOf(c);
        const nx = tr && tr.nextElementSibling;
        if (nx && nx.cells && nx.cells[idx]) {
          const v = blockText(nx.cells[idx]);
          if (v) return v;
        }
      }
      // (b) 同一格內就是「病史：內容」
      for (const c of cells) {
        const flat = txt(c);
        if (flat.startsWith(label) && flat.length > label.length + 2) {
          const full = blockText(c);
          const i = full.indexOf(label);
          return cleanML(full.slice(i + label.length).replace(/^[:：]\s*/, ''));
        }
      }
      // (c) textarea / input 的 name 或 id 含關鍵字
      for (const e of doc.querySelectorAll('textarea, input[type=text], input:not([type])')) {
        const key = `${e.name || ''} ${e.id || ''}`.toLowerCase();
        if (/hist|reason|purpose|content|memo/.test(key)) {
          const v = blockText(e);
          if (v) return v;
        }
      }
    }
    return '';
  };

  // ---------- 5. 診斷模式 ----------
  if (CONFIG.diagnose) {
    const url = linkOf(items[0].tr);
    console.log('[診斷] 明細頁網址：', url);
    if (!url) {
      console.error('[診斷] 找不到連結。這一列的 HTML：');
      console.log(items[0].tr.outerHTML.slice(0, 2000));
      return;
    }
    const doc = new DOMParser().parseFromString(
      await (await fetch(url, { credentials: 'include' })).text(),
      'text/html'
    );
    const pairs = [];
    for (const c of doc.querySelectorAll('td, th')) {
      const t = txt(c);
      if (t && t.length <= 12) {
        let n = c.nextElementSibling;
        while (n && !txt(n)) n = n.nextElementSibling;
        pairs.push({ 標題: t, 內容: (n ? txt(n) : '').slice(0, 120) });
      }
    }
    console.log('[診斷] 明細頁「標題 → 內容」候選（前 40 組）：');
    console.table(pairs.slice(0, 40));
    const got = extractField(doc, CONFIG.detailLabels);
    console.log('[診斷] 目前設定抓到：', JSON.stringify(got));

    // 傾印「未經處理」的原始病史，用來檢查斷行還原是否正確
    const rawEl = [...doc.querySelectorAll('td, th')].find(
      (c) => txt(c).replace(/[:：\s]/g, '') === CONFIG.detailLabels[0]
    );
    let rawNext = rawEl && rawEl.nextElementSibling;
    while (rawNext && !txt(rawNext)) rawNext = rawNext.nextElementSibling;
    if (rawNext) {
      const ta = rawNext.querySelector('textarea, input[type=text], input:not([type])');
      const raw = ta
        ? ta.value || ta.textContent || ''
        : rawNext.innerHTML.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
      const rawLines = raw.replace(/\r\n?/g, '\n').split('\n');
      console.log('[診斷] 原始病史共', rawLines.length, '行，各行長度：', rawLines.map((l) => l.length));
      console.log('[診斷] 原始病史（含行尾標記 ¶，方便看出行尾空白）：');
      console.log(rawLines.map((l, i) => String(i).padStart(2) + '|' + l + '¶').join('\n'));
      window.__PMR_RAW__ = raw;
      console.log('[診斷] 原始文字已存在 window.__PMR_RAW__');
    }

    window.__PMR_DOC__ = doc;
    return;
  }

  // ---------- 6. 逐筆抓明細 ----------
  if (CONFIG.fetchDetail) {
    const cache = new Map();
    let ok = 0,
      fail = 0;
    for (let i = 0; i < items.length; i++) {
      const { obj, tr } = items[i];
      const url = linkOf(tr);
      if (!url) {
        fail++;
        continue;
      }
      try {
        let doc = cache.get(url);
        if (!doc) {
          const res = await fetch(url, { credentials: 'include' });
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          doc = new DOMParser().parseFromString(await res.text(), 'text/html');
          cache.set(url, doc);
        }
        let v = extractField(doc, CONFIG.detailLabels);
        if (!CONFIG.preserveLineBreaks) v = clean(v);
        obj['會診內容'] = v;
        v ? ok++ : fail++;
      } catch (e) {
        console.warn(`[PMR] 第 ${i + 1} 筆失敗：`, e.message);
        fail++;
      }
      if ((i + 1) % 5 === 0 || i === items.length - 1)
        console.log(`[PMR] 進度 ${i + 1}/${items.length}（成功 ${ok}、無內容 ${fail}）`);
      if (i < items.length - 1) await sleep(CONFIG.delayMs);
    }
    if (ok === 0)
      console.warn('[PMR] 完全沒抓到會診內容。請把 diagnose 設為 true 跑一次看明細頁結構。');
  }

  // ---------- 7. 排序 ----------
  const rows = items.map((x) => x.obj);
  if (CONFIG.sortBy && rows.length && CONFIG.sortBy in rows[0]) {
    rows.sort((a, b) => {
      const x = a[CONFIG.sortBy] || '';
      const y = b[CONFIG.sortBy] || '';
      return CONFIG.sortDesc ? y.localeCompare(x) : x.localeCompare(y);
    });
  }
  console.table(rows.slice(0, 10));

  // ---------- 8. 產生並下載 xlsx ----------
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(
    d.getDate()
  ).padStart(2, '0')}`;
  const name = `PMR_${stamp}`;

  const bytes = buildXlsx(CONFIG.columns, rows, name);
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${name}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);

  console.log(`%c[PMR] ✓ 已下載 ${name}.xlsx（${rows.length} 筆）`, 'color:green;font-weight:bold');
  window.__PMR__ = { headers: CONFIG.columns.map((c) => c[0]), rows };
})();