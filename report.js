/* 每月抽查報告自動更新 —— 瀏覽器版（純前端，直接操作 docx 內部 OOXML 格式，不需伺服器）。
   跟 update_audit_report.py 是同一套邏輯的兩份實作（Python 版先驗證過真實檔案可行），
   這裡刻意逐一對照那份腳本的函式與規則，任何一邊修規則，另一邊也要記得同步改。 */
(function () {
  "use strict";

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
  const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const CASE_ID_RE = /^\d{13}$/;
  const TRAILING_ROAD_RE = /\(([^()]+)\)\s*$/;
  const LS_CASES = "dispatch_cases_v1";
  const LS_LASTMONTH_FULL = "dispatch_lastmonth_full_v1";
  const LS_LASTMONTH_AUDIT = "dispatch_lastmonth_audit_v1";

  // File System Access API 只有安全來源（https 或 localhost）才能用，Firefox 完全不支援，
  // 用 file:// 直接開啟這個工具也不算安全來源——都會退回「選檔案 + 觸發下載」的舊做法
  const supportsFsAccess = typeof window.showDirectoryPicker === "function" && window.isSecureContext;
  let reportDirHandle = null;      // 選定的報告資料夾，之後產生報告直接寫進去，不用每次跳出存檔視窗
  let reportTemplateHandle = null; // 資料夾裡自動抓到（或使用者從下拉選單挑）的月報範本檔案 handle

  // 案件項目 -> 查證內容中間那句觀察描述的固定說法，跟 update_audit_report.py 的
  // CATEGORY_PHRASES 是同一份對照表（從 D:\派工抽查 底下 18 份過去幾個月的真實抽查報告，
  // 撈出將近 900 筆歷史「查證內容」欄位文字整理出來，不是憑空編的）。兩邊要維持一致，
  // 一邊改規則另一邊也要記得同步改。
  const CATEGORY_PHRASES = {
    "市區道路坑洞處理": "道路坑洞已臨補",
    "交通號誌異常": "行車號誌已正常運作",
    "路燈不亮或損壞": "路燈已正常放亮",
    "路樹處理": "路樹已修剪",
    "交通標誌及設施物損壞(含汙損)、傾斜": "反光鏡已調正",
    "道路側溝溝蓋(含周邊)損壞遺失": "側溝溝蓋已修復固定",
    "道路散落物或油漬處理": "路面油漬及散落物已清除",
    "用戶無水、漏水報修": "該址已無漏水情事",
    "交通號誌電纜線垂落及設施損壞": "號誌線路及設施已修復正常",
    "鄰里無主垃圾清運": "該址已無垃圾",
    "人孔蓋(含周邊)破損、遺失處理": "人孔蓋已修復固定",
    "雨水下水道側溝清淤": "側溝已清疏",
  };

  // 有些案件項目底下其實包了好幾種完全不同的實際狀況（例如「路燈不亮或損壞」同時涵蓋不亮、
  // 電線垂落、燈桿傾斜、零件脫落…），只看項目分類套一句固定說法會講錯（使用者實測抓到：北投一筆
  // 「路燈不亮或損壞」案件實際是燈桿傾斜，套用「路燈已正常放亮」文不對題）。改成先比對案件原始
  // 內容（c.content）裡有沒有更具體的關鍵字，命中才用對應的說法；比對不到才退回類別預設說法，
  // 這種情況會回傳 confident:false，畫面上要標示「請核對」提醒使用者這句沒有比對案件內容、
  // 只是用項目分類猜的，務必自己確認。跟 update_audit_report.py 的同名函式要維持一致。
  const CONTENT_KEYWORD_PHRASES = {
    "路燈不亮或損壞": [
      [/傾斜|歪斜|歪掉/, "燈桿已扶正"],
      [/電線|纜線|垂落/, "路燈電線已無垂落"],
      [/外蓋|燈罩|零件.*脫落|鬆脫/, "燈桿外蓋已固定"],
      [/基座|底座/, "路燈基座已修復"],
      [/不亮|不放光|熄|故障|閃爍/, "路燈已正常放亮"],
    ],
    "交通標誌及設施物損壞(含汙損)、傾斜": [
      [/反射鏡|反光鏡/, "反光鏡已調正"],
      [/汙損|髒污|塗鴉/, "標誌牌汙損已清潔"],
      [/傾斜|歪斜|歪掉/, "標誌已扶正"],
      [/斷裂|損壞|毀損/, "已更換固定"],
    ],
    "交通號誌電纜線垂落及設施損壞": [
      [/垂落|電纜線/, "號誌電纜線已妥善固定，無垂落情形"],
      [/傾斜|歪斜/, "號誌燈桿已扶正"],
      [/按鈕|按壓/, "行人按鈕已修復正常運作"],
    ],
    "道路側溝溝蓋(含周邊)損壞遺失": [
      [/遺失|不見/, "側溝溝蓋已補齊固定"],
      [/鬆動|聲響/, "側溝溝蓋已固定，無鬆動聲響"],
      [/凹陷|破損/, "側溝溝蓋周邊已修復"],
    ],
    "人孔蓋(含周邊)破損、遺失處理": [
      [/突起|不平|沒有放平|凸起|翹起/, "人孔蓋已回復平整"],
      [/遺失|不見|沒有蓋/, "人孔蓋已補齊固定"],
      [/鬆動|聲響|搖晃/, "人孔蓋已固定，無鬆動聲響"],
      [/破損|裂/, "人孔蓋已修復完成"],
    ],
  };

  function buildVerificationPhrase(category, content) {
    const normalized = (category || "").replace(/\s+/g, "");
    let matchedKey = null;
    for (const key in CATEGORY_PHRASES) {
      if (key.replace(/\s+/g, "") === normalized) { matchedKey = key; break; }
    }
    const rules = matchedKey ? CONTENT_KEYWORD_PHRASES[matchedKey] : null;
    if (rules) {
      for (const [pattern, phrase] of rules) {
        if (pattern.test(content || "")) return { phrase, confident: true };
      }
    }
    const fallback = matchedKey ? CATEGORY_PHRASES[matchedKey] : (category ? `${category}已完成改善` : "現場已完成改善");
    // 這個項目根本沒有訂細分規則（rules 是 null）時，就是原本的行為，不用特別標記；
    // 有訂規則、但案件內容比對不到任何一條，才是真的沒把握，要提醒使用者自己核對
    return { phrase: fallback, confident: !rules };
  }

  function directChildren(el, localName) {
    return Array.from(el.children).filter((c) => c.localName === localName);
  }
  function descendantsNS(el, ns, localName) {
    return Array.from(el.getElementsByTagNameNS(ns, localName));
  }
  function paragraphText(p) {
    return descendantsNS(p, W_NS, "t").map((t) => t.textContent || "").join("");
  }

  // ---------- 地址簡化（best-effort，使用者確認會自己在最後報告手動微調，不用做到完美規則） ----------
  // 規則跟 update_audit_report.py 的 simplify_address() 完全一致：
  //   1. 有「里/村」+ 結尾「(道路名)」括號的話，取最後一個「里/村」之後、括號之前那一段
  //      （真實資料顯示這段幾乎都已經是乾淨的「路名+號」）。
  //   2. 否則單純去掉開頭重複的「台北市/臺北市」。
  //   3. 結果如果還殘留英文字母／市／區（後面沒接「號」）等雜訊，標記 lowConfidence，
  //      提示使用者這幾筆要特別檢查。
  function simplifyAddress(raw) {
    raw = (raw || "").trim();
    const trailingMatch = raw.match(TRAILING_ROAD_RE);
    const villageMatches = Array.from(raw.matchAll(/[里村]/g));
    let candidate;
    if (trailingMatch && villageMatches.length) {
      const last = villageMatches[villageMatches.length - 1];
      const start = last.index + last[0].length;
      const end = trailingMatch.index;
      candidate = raw.slice(start, end).trim();
    } else {
      candidate = raw.replace(/^(?:台北市|臺北市)+/, "").trim();
    }
    const lowConfidence = /[A-Za-z]|市|區(?!.*號)/.test(candidate) || !candidate;
    return { text: candidate, lowConfidence };
  }

  // ---------- 直接讀工具本身已匯入的案件資料（本月 state.cases／上月抽查 getLastMonthCases()），
  // 不需要另外匯出/選一份查訪路線 docx 當中介——查訪路線文件的「地址」欄位本來就只是原封不動印出
  // c.address 加「台北市」前綴而已（見 app.js 匯出 Word 那段），並沒有比案件資料本身更多的資訊，
  // 繞一圈用 Word 文字反解析（regex 對付排版、換行）純粹是多一層失真風險，不如直接讀。
  // report.js 是獨立的 IIFE，不能直接呼叫 app.js 內部的 state／getLastMonthCases()（不同函式作用域），
  // 所以這裡直接讀同一組 localStorage key、自己重算一次上月合併邏輯（跟 app.js 的 getLastMonthCases()
  // 對照：這裡用未加「抽」前綴的原始 id，因為要拿去跟照片檔名比對，過濾掉字首前綴的問題）
  function loadAppCaseData() {
    let thisMonth = [];
    try { thisMonth = JSON.parse(localStorage.getItem(LS_CASES) || "[]"); } catch (e) {}
    let fullList = [], auditList = [];
    try { fullList = JSON.parse(localStorage.getItem(LS_LASTMONTH_FULL) || "[]"); } catch (e) {}
    try { auditList = JSON.parse(localStorage.getItem(LS_LASTMONTH_AUDIT) || "[]"); } catch (e) {}
    const fullById = new Map(fullList.map((c) => [c.id, c]));
    const lastMonth = auditList.map((a) => {
      const full = fullById.get(a.id);
      return {
        id: a.id,
        category: a.category || (full && full.category) || "",
        district: a.district || (full && full.district) || "",
        agency: full ? full.agency : "",
        address: full ? full.address : "",
        content: full ? full.content : "",
      };
    });
    return { thisMonth, lastMonth };
  }

  // ---------- 照片 EXIF 拍攝日期（手動解析 JPEG APP1/Exif 區段，只抓 DateTimeOriginal/DateTime） ----------
  function readExifDate(arrayBuffer) {
    const view = new DataView(arrayBuffer);
    if (view.byteLength < 4 || view.getUint16(0) !== 0xffd8) return null; // 不是 JPEG
    let offset = 2;
    while (offset + 4 <= view.byteLength) {
      const marker = view.getUint16(offset);
      if (marker === 0xffd9 || marker === 0xffda) break; // EOI / SOS，後面不會再有 metadata
      const size = view.getUint16(offset + 2);
      if (marker === 0xffe1) {
        const exifStart = offset + 4;
        if (exifStart + 6 <= view.byteLength) {
          const tagBytes = new Uint8Array(arrayBuffer, exifStart, 6);
          const tag = String.fromCharCode.apply(null, tagBytes);
          if (tag.startsWith("Exif")) {
            const dt = readExifDateFromTiff(view, exifStart + 6, arrayBuffer);
            if (dt) return dt;
          }
        }
      }
      offset += 2 + size;
    }
    return null;
  }

  function readExifDateFromTiff(view, tiffStart, arrayBuffer) {
    const little = view.getUint16(tiffStart) === 0x4949; // 'II'
    const u16 = (o) => view.getUint16(o, little);
    const u32 = (o) => view.getUint32(o, little);
    const firstIfdOffset = u32(tiffStart + 4);

    function readIfd(ifdStart) {
      let dtOriginal = null, dtFallback = null;
      const numEntries = u16(ifdStart);
      for (let i = 0; i < numEntries; i++) {
        const entryOffset = ifdStart + 2 + i * 12;
        const tagId = u16(entryOffset);
        const type = u16(entryOffset + 2);
        const count = u32(entryOffset + 4);
        const valueOffset = entryOffset + 8;
        if (tagId === 0x9003 || tagId === 0x0132) {
          let strStart = (type === 2 && count > 4) ? tiffStart + u32(valueOffset) : valueOffset;
          const bytes = new Uint8Array(arrayBuffer, strStart, count);
          let str = String.fromCharCode.apply(null, bytes);
          str = str.replace(/\0.*$/, "");
          if (tagId === 0x9003) dtOriginal = str; else dtFallback = str;
        }
        if (tagId === 0x8769) { // Exif SubIFD 指標，DateTimeOriginal 通常在這裡面
          const subIfdStart = tiffStart + u32(valueOffset);
          const sub = readIfd(subIfdStart);
          if (sub) dtOriginal = dtOriginal || sub;
        }
      }
      return dtOriginal || dtFallback;
    }

    const dtStr = readIfd(tiffStart + firstIfdOffset);
    if (!dtStr) return null;
    const m = dtStr.match(/(\d{4}):(\d{2}):(\d{2})/);
    return m ? { y: +m[1], m: +m[2], d: +m[3] } : null;
  }

  async function photoExifDate(file) {
    try {
      const buf = await file.arrayBuffer();
      return readExifDate(buf);
    } catch (e) {
      return null;
    }
  }

  // ---------- 在月報範本裡找到某行政區的表格（標題段落文字含「（{district}區）」，後面第一個表格） ----------
  function findDistrictTable(bodyEl, district) {
    let headingFound = false;
    for (const child of Array.from(bodyEl.children)) {
      if (child.localName === "p") {
        if (paragraphText(child).includes(`（${district}區）`)) { headingFound = true; continue; }
      } else if (headingFound && child.localName === "tbl") {
        return child;
      }
    }
    return null;
  }

  function tableRows(tbl) { return directChildren(tbl, "tr"); }
  function rowCells(tr) { return directChildren(tr, "tc"); }

  function setRunText(run, text) {
    const ts = directChildren(run, "t");
    if (ts.length) {
      ts[0].textContent = text;
      ts[0].setAttribute("xml:space", "preserve");
      for (let j = 1; j < ts.length; j++) ts[j].textContent = "";
    } else {
      const t = run.ownerDocument.createElementNS(W_NS, "w:t");
      t.setAttribute("xml:space", "preserve");
      t.textContent = text;
      run.appendChild(t);
    }
  }

  // 把 cell 內容換成 texts（一個 paragraph 一行）。盡量重用原有段落的第一個 run 來設定文字，
  // 藉此保留原本的字型/樣式，而不是新增一個沒有格式的 run —— 跟 Python 版 clear_and_set_text() 同邏輯
  function setCellTexts(tc, texts) {
    const paragraphs = directChildren(tc, "p");
    texts.forEach((text, i) => {
      if (i < paragraphs.length) {
        const p = paragraphs[i];
        const runs = directChildren(p, "r");
        if (runs.length) {
          setRunText(runs[0], text);
          for (let j = 1; j < runs.length; j++) setRunText(runs[j], "");
        } else {
          const r = tc.ownerDocument.createElementNS(W_NS, "w:r");
          setRunText(r, text);
          p.appendChild(r);
        }
      } else if (paragraphs.length) {
        const clone = paragraphs[paragraphs.length - 1].cloneNode(true);
        const cloneRuns = directChildren(clone, "r");
        if (cloneRuns.length) {
          setRunText(cloneRuns[0], text);
          for (let j = 1; j < cloneRuns.length; j++) setRunText(cloneRuns[j], "");
        }
        tc.appendChild(clone);
      }
    });
    for (let i = texts.length; i < paragraphs.length; i++) {
      directChildren(paragraphs[i], "r").forEach((r) => setRunText(r, ""));
    }
  }

  // ---------- 照片縮圖（canvas，避免手機拍的原始照片幾MB直接塞進 docx 越改越肥） ----------
  function resizeImageToJpegBlob(file, maxDim, quality) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        let { width, height } = img;
        const scale = Math.min(1, maxDim / Math.max(width, height));
        width = Math.round(width * scale);
        height = Math.round(height * scale);
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        canvas.toBlob((blob) => {
          URL.revokeObjectURL(url);
          blob ? resolve(blob) : reject(new Error("縮圖失敗"));
        }, "image/jpeg", quality);
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("讀取照片失敗")); };
      img.src = url;
    });
  }

  // 瀏覽器的 XMLSerializer 序列化一個原本就帶 XML 宣告的 Document 時，本身就會把宣告一起印出來
  // （跟 python-docx/一般認知不同，這裡曾經因為手動再補一次宣告、變成兩個宣告而讓 Word 檔壞掉），
  // 這裡只在序列化結果真的沒有宣告時才手動補上，避免重複
  function xmlToString(doc) {
    const s = new XMLSerializer().serializeToString(doc);
    return s.startsWith("<?xml") ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
  }

  async function nextMediaIndex(zip) {
    let maxIdx = 0;
    Object.keys(zip.files).forEach((f) => {
      const m = f.match(/^word\/media\/image(\d+)\.\w+$/);
      if (m) maxIdx = Math.max(maxIdx, parseInt(m[1], 10));
    });
    return maxIdx + 1;
  }

  async function addImagePart(zip, blob) {
    const idx = await nextMediaIndex(zip);
    const mediaPath = `word/media/image${idx}.jpeg`;
    zip.file(mediaPath, blob);

    const relsPath = "word/_rels/document.xml.rels";
    const relsXmlStr = await zip.file(relsPath).async("string");
    const relsDoc = new DOMParser().parseFromString(relsXmlStr, "application/xml");
    let maxRid = 0;
    Array.from(relsDoc.getElementsByTagNameNS(RELS_NS, "Relationship")).forEach((r) => {
      const m = (r.getAttribute("Id") || "").match(/^rId(\d+)$/);
      if (m) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    });
    const newRid = `rId${maxRid + 1}`;
    const relEl = relsDoc.createElementNS(RELS_NS, "Relationship");
    relEl.setAttribute("Id", newRid);
    relEl.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    relEl.setAttribute("Target", `media/image${idx}.jpeg`);
    relsDoc.documentElement.appendChild(relEl);
    zip.file(relsPath, xmlToString(relsDoc));

    const ctPath = "[Content_Types].xml";
    const ctXmlStr = await zip.file(ctPath).async("string");
    const ctDoc = new DOMParser().parseFromString(ctXmlStr, "application/xml");
    const hasJpeg = Array.from(ctDoc.getElementsByTagNameNS(CT_NS, "Default")).some((d) => d.getAttribute("Extension") === "jpeg");
    if (!hasJpeg) {
      const d = ctDoc.createElementNS(CT_NS, "Default");
      d.setAttribute("Extension", "jpeg");
      d.setAttribute("ContentType", "image/jpeg");
      ctDoc.documentElement.appendChild(d);
      zip.file(ctPath, xmlToString(ctDoc));
    }
    return newRid;
  }

  function buildDrawingRunXml(rId, picId, cx, cy) {
    return `<w:r xmlns:w="${W_NS}" xmlns:wp="${WP_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${R_NS}">` +
      `<w:rPr><w:noProof/></w:rPr>` +
      `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${picId}" name="Picture ${picId}"/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  // cx/cy 固定 5.5cm x 4.1cm（1cm = 360000 EMU），跟範本裡其他已填好的照片尺寸一致
  const PHOTO_CX = 1980000, PHOTO_CY = 1476000;
  let picIdCounter = 90000;

  async function replaceCellImage(tc, resizedBlob, mainDoc, zip) {
    directChildren(tc, "p").forEach((p) => {
      directChildren(p, "r").forEach((r) => p.removeChild(r));
    });
    const rId = await addImagePart(zip, resizedBlob);
    const runXml = buildDrawingRunXml(rId, ++picIdCounter, PHOTO_CX, PHOTO_CY);
    const parsed = new DOMParser().parseFromString(runXml, "application/xml");
    const imported = mainDoc.importNode(parsed.documentElement, true);
    const targetP = directChildren(tc, "p")[0];
    targetP.appendChild(imported);
  }

  // ---------- 主流程：跟 update_audit_report.py 的 build_report() 逐項對應 ----------
  async function buildReport({ reportFile, photoFiles, district, dateOverride, dryRun }) {
    const fallbackDate = dateOverride || null;

    const photoByStem = {};
    for (const f of photoFiles) {
      const stem = f.name.replace(/\.(jpe?g)$/i, "");
      if (CASE_ID_RE.test(stem)) photoByStem[stem] = f;
    }

    const { thisMonth, lastMonth } = loadAppCaseData();
    const districtName = `${district}區`;
    const fresh = thisMonth
      .filter((c) => c.district === districtName && photoByStem[c.id])
      .map((c) => ({ ...c, isLastMonth: false }));
    const recheck = lastMonth
      .filter((c) => c.district === districtName && photoByStem[c.id])
      .map((c) => ({ ...c, isLastMonth: true }));
    // 使用者要求錯誤訊息不用解釋原因、不用給建議，一句話講完就好
    const matched = fresh.concat(recheck);
    if (matched.length !== 5) {
      throw new ReportError(
        matched.length === 0
          ? "你選擇的照片找不到對應的案件編號。"
          : `符合的案件只有 ${matched.length} 筆，需要剛好 5 筆（${matched.map((s) => s.id).join("、")}）。`
      );
    }
    if (recheck.length !== 1 || fresh.length !== 4) {
      throw new ReportError(`需要剛好 4 筆本月＋1 筆上月，目前是 ${fresh.length} 筆本月＋${recheck.length} 筆上月。`);
    }
    fresh.sort((a, b) => a.id.localeCompare(b.id));
    const rowsData = fresh.concat(recheck); // 第1~4列＝本月抽查，第5列＝上月複查

    const reportBuf = await reportFile.arrayBuffer();
    const zip = await JSZip.loadAsync(reportBuf);
    const docXmlStr = await zip.file("word/document.xml").async("string");
    const mainDoc = new DOMParser().parseFromString(docXmlStr, "application/xml");
    const body = mainDoc.getElementsByTagNameNS(W_NS, "body")[0];
    const table = findDistrictTable(body, district);
    if (!table) throw new ReportError(`在範本裡找不到「（${district}區）」的標題／表格，請確認行政區名稱跟範本裡的一致。`);
    const rows = tableRows(table);
    if (rows.length !== 6) throw new ReportError(`找到的表格不是預期的 6 列（1 標題+5 案件），實際是 ${rows.length} 列，不敢自動改，需要人工檢查範本結構是否變過。`);

    const rowDates = {};
    const distinctDates = new Set();
    for (const s of rowsData) {
      const exif = await photoExifDate(photoByStem[s.id]);
      const md = exif ? `${exif.m}/${exif.d}` : fallbackDate;
      if (!md) throw new ReportError(`案號 ${s.id} 的照片沒有 EXIF 拍攝日期，且沒有指定備援日期，無法決定查證日期。`);
      rowDates[s.id] = md;
      distinctDates.add(md);
    }

    const preview = [];
    for (let i = 0; i < rowsData.length; i++) {
      const s = rowsData[i];
      const { text: addr, lowConfidence } = simplifyAddress(s.address);
      const stage = s.isLastMonth ? "複查" : "抽查";
      const closing = s.isLastMonth ? "，權責機關確已完成案件處理及抽查作業。" : "，權責機關確依限完成案件處理作業。";
      const { phrase: verificationPhrase, confident: phraseConfident } = buildVerificationPhrase(s.category, s.content);
      const contentLine2 = `經實地查證，${verificationPhrase}` + closing;
      const dateMd = rowDates[s.id];

      preview.push({
        row: i + 1, id: s.id, stage, category: s.category, agency: s.agency,
        date: dateMd, address: addr, addressLowConfidence: lowConfidence,
        addressRaw: s.address, photoName: photoByStem[s.id].name,
        verificationPhrase, phraseLowConfidence: !phraseConfident,
      });

      if (dryRun) continue;

      const cells = rowCells(rows[i + 1]);
      setCellTexts(cells[1], [s.id]);
      setCellTexts(cells[2], [dateMd]);
      setCellTexts(cells[3], [s.category]);
      setCellTexts(cells[4], [s.agency]);
      // cells[5]（查證階段）跟 cells[6]（是否通過）是跟「列位置」綁定的既有樣板文字，維持原樣
      setCellTexts(cells[7], [`案址：${addr}`, contentLine2]);
      const resized = await resizeImageToJpegBlob(photoByStem[s.id], 1200, 0.85);
      await replaceCellImage(cells[8], resized, mainDoc, zip);
    }

    const result = { preview, distinctDates: Array.from(distinctDates) };
    if (dryRun) return result;

    zip.file("word/document.xml", xmlToString(mainDoc));
    result.blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
    return result;
  }

  function ReportError(message) { this.message = message; this.name = "ReportError"; }
  ReportError.prototype = Object.create(Error.prototype);

  window.AuditReportModule = { buildReport, simplifyAddress, loadAppCaseData, photoExifDate, ReportError };

  // ---------- UI ----------
  const KNOWN_DISTRICTS = ["中正", "萬華", "松山", "信義", "大安", "文山", "中山", "大同", "士林", "北投", "內湖", "南港"];

  function guessDistrict(text) {
    return KNOWN_DISTRICTS.find((d) => text.includes(d)) || "";
  }

  function pad2(n) { return String(n).padStart(2, "0"); }

  function renderPreviewTable(preview) {
    const rows = preview.map((r) => `
      <tr>
        <td>${r.row}</td>
        <td>${r.stage}</td>
        <td>${escapeHtmlLocal(r.id)}</td>
        <td>${escapeHtmlLocal(r.date)}</td>
        <td>${escapeHtmlLocal(r.category)}</td>
        <td>${escapeHtmlLocal(r.agency)}</td>
        <td>${escapeHtmlLocal(r.address)}${r.addressLowConfidence ? ' <span style="color:#c9821a">⚠ 請確認</span>' : ""}</td>
        <td>${escapeHtmlLocal(r.verificationPhrase)}${r.phraseLowConfidence ? ' <span style="color:#c9821a">⚠ 請核對</span>' : ""}</td>
        <td>${escapeHtmlLocal(r.photoName)}</td>
      </tr>`).join("");
    return `<table class="case-table">
      <thead><tr><th>列</th><th>階段</th><th>案號</th><th>日期</th><th>項目</th><th>機關</th><th>簡化地址</th><th>查證說法</th><th>照片</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  function escapeHtmlLocal(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // 輸出檔名固定格式「{原始檔名}_已更新_{行政區}_{YYYYMMDD}_{HHMM}.docx」，從這裡面把時間戳記
  // 抓出來（抓不到就回傳 null，代表這不是一份「已更新」檔案，是原始範本）
  function extractUpdatedTimestamp(name) {
    const m = name.match(/_已更新_[^_]+_(\d{8})_(\d{4})\.docx$/i);
    return m ? m[1] + m[2] : null;
  }
  // 拿掉「_已更新_{行政區}_{時間戳記}」尾巴，回推出乾淨的原始檔名——這樣不管接續哪一份「已更新」
  // 檔案繼續產生，輸出檔名都還是用原始範本的名字加新時間戳記，不會一次一次疊字越疊越長
  function deriveBaseName(name) {
    return name.replace(/_已更新_[^_]+_\d{8}_\d{4}(\.docx)$/i, "$1");
  }

  // 選好資料夾後：申請寫入權限（存回同一個資料夾要用），掃資料夾裡的 docx 找範本，排除 Word 開啟中
  // 的暫存鎖定檔「~$」。使用者反映：先做士林產生一份「已更新」檔案，接著做北投時，如果又抓回最原始
  // 的範本，兩個行政區的更新不會疊在同一份檔案裡——改成「已更新」檔案不再排除，而是優先接續其中
  // 時間戳記最新的那份繼續疊加；完全沒有「已更新」檔案時才用原始範本。剛好只有 1 個候選就直接用；
  // 有多個（例如同時有原始範本＋好幾份不同時間點的已更新檔案）就列出來、預設選最新的，也能自己換
  async function pickReportFolder() {
    const statusEl = document.getElementById("reportFileStatus");
    const pickSel = document.getElementById("reportTemplatePick");
    let dir;
    try {
      dir = await window.showDirectoryPicker();
    } catch (e) {
      return; // 使用者按取消（AbortError），不用特別提示
    }
    let perm = await dir.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await dir.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") {
      statusEl.textContent = "沒有取得這個資料夾的寫入權限，沒辦法把報告存進去。";
      return;
    }
    reportDirHandle = dir;

    const candidates = [];
    for await (const [name, handle] of dir.entries()) {
      if (handle.kind === "file" && /\.docx$/i.test(name) && !name.startsWith("~$")) {
        candidates.push({ name, handle, ts: extractUpdatedTimestamp(name) });
      }
    }
    if (!candidates.length) {
      statusEl.textContent = `「${dir.name}」資料夾裡找不到任何 docx 檔案。`;
      reportTemplateHandle = null;
      pickSel.style.display = "none";
      return;
    }
    // 「已更新」檔案（有時間戳記）優先，且時間戳記新的排前面；原始範本（沒有時間戳記）排最後
    candidates.sort((a, b) => {
      if (a.ts && b.ts) return b.ts.localeCompare(a.ts);
      if (a.ts) return -1;
      if (b.ts) return 1;
      return 0;
    });

    if (candidates.length === 1) {
      reportTemplateHandle = candidates[0].handle;
      pickSel.style.display = "none";
      statusEl.textContent = `已選擇資料夾「${dir.name}」，範本：${candidates[0].name}`;
    } else {
      pickSel.innerHTML = "";
      candidates.forEach((c, i) => {
        const opt = document.createElement("option");
        opt.value = c.name; opt.textContent = c.name + (i === 0 ? "（最新，預設）" : "");
        pickSel.appendChild(opt);
      });
      pickSel.value = candidates[0].name;
      pickSel.style.display = "";
      reportTemplateHandle = candidates[0].handle;
      statusEl.textContent = `「${dir.name}」資料夾裡有 ${candidates.length} 份 docx，已預設選最新的一份，也可以自己換：`;
      pickSel.onchange = () => {
        const found = candidates.find((c) => c.name === pickSel.value);
        reportTemplateHandle = found ? found.handle : null;
      };
    }
  }

  async function getReportFile() {
    if (supportsFsAccess && reportTemplateHandle) return reportTemplateHandle.getFile();
    const fallbackInput = document.getElementById("reportFileInputFallback");
    return (fallbackInput && fallbackInput.files[0]) || null;
  }

  async function runBuild(dryRun) {
    const statusEl = document.getElementById("reportStatus");
    const previewEl = document.getElementById("reportPreview");
    const photosInput = document.getElementById("reportPhotosInput");
    const districtSel = document.getElementById("reportDistrict");
    const dateInput = document.getElementById("reportDateOverride");

    const reportFile = await getReportFile();
    if (!reportFile) { alert(supportsFsAccess ? "請先選擇報告資料夾。" : "請先選擇月報範本 docx。"); return; }
    const photoFiles = Array.from(photosInput.files || []).filter((f) => /\.jpe?g$/i.test(f.name));
    if (!photoFiles.length) { alert("請先選擇照片資料夾（裡面要有以案件編號命名的 jpg 照片）。"); return; }
    if (!districtSel.value) { alert("請選擇行政區。"); return; }

    statusEl.textContent = "處理中…";
    previewEl.innerHTML = "";
    try {
      const result = await window.AuditReportModule.buildReport({
        reportFile,
        photoFiles,
        district: districtSel.value,
        dateOverride: dateInput.value.trim() || null,
        dryRun,
      });
      let warn = "";
      if (result.distinctDates.length > 1) {
        warn = `<div class="hint" style="color:#c9821a;margin-top:6px">⚠ 這 5 張照片拍攝日期不只一天：${result.distinctDates.join("、")}，已各自照實際拍攝日期填入，請確認這是預期中的狀況。</div>`;
      }
      previewEl.innerHTML = renderPreviewTable(result.preview) + warn;
      if (dryRun) {
        statusEl.textContent = "預覽完成（尚未產生檔案，確認沒問題後按下方按鈕產生）";
      } else if (result.blob) {
        const saved = await saveResult(result.blob, reportFile.name, districtSel.value);
        statusEl.textContent = saved.mode === "folder"
          ? `已存到「${reportDirHandle.name}」資料夾：${saved.filename}`
          : `已下載：${saved.filename}`;
      }
    } catch (e) {
      statusEl.textContent = "";
      if (e instanceof window.AuditReportModule.ReportError) {
        previewEl.innerHTML = `<div class="empty-state">⚠ ${escapeHtmlLocal(e.message)}</div>`;
      } else {
        previewEl.innerHTML = `<div class="empty-state">發生錯誤：${escapeHtmlLocal(e.message || String(e))}</div>`;
        console.error(e);
      }
    }
  }

  // 有 File System Access API 可用時（Chrome/Edge，且是 https 或 localhost 這種「安全來源」），
  // 拿模板所在資料夾的 handle 直接寫檔進去，不用每次都跳出「另存新檔」視窗，也不會亂跑去下載資料夾；
  // 不支援時（Firefox、或用 file:// 直接開啟這個工具）退回原本「觸發瀏覽器下載」的做法
  async function saveResult(blob, originalName, district) {
    const now = new Date();
    const ts = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;
    // originalName 可能本身就是一份「已更新」檔案（接續上次的結果繼續疊加），先拿掉那段尾巴
    // 回推乾淨的原始檔名，輸出檔名才不會一次一次疊字（...已更新_士林_ts1_已更新_北投_ts2...）
    const base = deriveBaseName(originalName).replace(/\.docx$/i, "");
    const filename = `${base}_已更新_${district}_${ts}.docx`;

    if (supportsFsAccess && reportDirHandle) {
      const fileHandle = await reportDirHandle.getFileHandle(filename, { create: true });
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      return { mode: "folder", filename };
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    return { mode: "download", filename };
  }

  document.addEventListener("DOMContentLoaded", () => {
    const photosInput = document.getElementById("reportPhotosInput");
    const districtSel = document.getElementById("reportDistrict");
    if (!photosInput || !districtSel) return; // 這個頁籤還沒建立在畫面上（理論上不會發生，防禦性檢查）

    KNOWN_DISTRICTS.forEach((d) => {
      const opt = document.createElement("option");
      opt.value = d; opt.textContent = d;
      districtSel.appendChild(opt);
    });

    // 「選擇照片資料夾」還是維持原本「隱藏的 <input type=file webkitdirectory>」做法（跟資料匯入
    // 頁籤的既有按鈕/input 配對方式一樣），只讀不寫，不需要 File System Access API
    function wireFileButton(btnId, inputId, statusId, describe) {
      const btn = document.getElementById(btnId);
      const input = document.getElementById(inputId);
      const status = document.getElementById(statusId);
      btn.addEventListener("click", () => input.click());
      input.addEventListener("change", () => { status.textContent = describe(input.files); });
    }
    wireFileButton("btnChooseReportPhotos", "reportPhotosInput", "reportPhotosStatus",
      (files) => {
        const jpgs = Array.from(files).filter((f) => /\.jpe?g$/i.test(f.name));
        return jpgs.length ? `已選擇資料夾，找到 ${jpgs.length} 張 jpg 照片` : "資料夾裡沒有找到 jpg 照片";
      });

    // 「選擇月報範本」這顆按鈕依瀏覽器支援度分兩種行為：
    //   支援 File System Access API → 選資料夾，自動抓範本、之後產生報告直接寫回同一個資料夾；
    //   不支援（Firefox／file:// 開啟）→ 退回選單一 docx 檔案、產生報告觸發瀏覽器下載這個舊做法
    const chooseReportBtn = document.getElementById("btnChooseReportFile");
    const genBtn = document.getElementById("btnReportGenerate");
    if (supportsFsAccess) {
      chooseReportBtn.textContent = "選擇報告資料夾";
      genBtn.textContent = "📄 產生報告（存到範本資料夾）";
      chooseReportBtn.addEventListener("click", pickReportFolder);
    } else {
      chooseReportBtn.textContent = "選擇月報範本 docx";
      const fallbackInput = document.getElementById("reportFileInputFallback");
      const fileStatus = document.getElementById("reportFileStatus");
      chooseReportBtn.addEventListener("click", () => fallbackInput.click());
      fallbackInput.addEventListener("change", () => {
        fileStatus.textContent = fallbackInput.files[0] ? `已選擇：${fallbackInput.files[0].name}` : "";
      });
    }

    // 沒有查訪路線檔名可以猜行政區了，改用照片資料夾的路徑（webkitdirectory 選資料夾時，每個檔案的
    // webkitRelativePath 會包含選到的資料夾名稱，例如「11508 士林/抽查照片/xxx.jpg」）
    photosInput.addEventListener("change", () => {
      const f = photosInput.files[0];
      if (f && f.webkitRelativePath) {
        const guess = guessDistrict(f.webkitRelativePath);
        if (guess) districtSel.value = guess;
      }
    });

    document.getElementById("btnReportPreview").addEventListener("click", () => runBuild(true));
    document.getElementById("btnReportGenerate").addEventListener("click", () => runBuild(false));
  });
})();
