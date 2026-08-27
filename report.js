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
  const STOP_HEADER_RE = /^\d+\.\s*案件編號：(\d{13})(（上月抽查）)?\s*承辦機關：([^\s　]+)/;
  const TRAILING_ROAD_RE = /\(([^()]+)\)\s*$/;

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

  function buildVerificationPhrase(category) {
    const normalized = (category || "").replace(/\s+/g, "");
    for (const key in CATEGORY_PHRASES) {
      if (key.replace(/\s+/g, "") === normalized) return CATEGORY_PHRASES[key];
    }
    return category ? `${category}已完成改善` : "現場已完成改善";
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

  // ---------- 路線文件解析（查訪路線_YYYYMMDD_HHMM{district}.docx，工具本身匯出的格式） ----------
  function parseRouteParagraphs(paragraphs) {
    const stops = [];
    let cur = null;
    for (const raw of paragraphs) {
      const line = raw.trim();
      if (!line) continue;
      const m = line.match(STOP_HEADER_RE);
      if (m) {
        if (cur) stops.push(cur);
        cur = {
          id: m[1], isLastMonth: !!m[2], agency: m[3],
          category: "", addressRaw: "", content: "", handling: "",
          order: stops.length + 1,
        };
        continue;
      }
      if (!cur) continue;
      if (line.startsWith("項目：")) cur.category = line.slice(3).trim();
      else if (line.startsWith("地址：")) cur.addressRaw = line.slice(3).trim();
      else if (line.startsWith("內容：")) cur.content = (cur.content + " " + line.slice(3).trim()).trim();
      else if (line.startsWith("處理：")) cur.handling = (cur.handling + " " + line.slice(3).trim()).trim();
      else if (cur.content && !cur.handling) cur.content += " " + line;
    }
    if (cur) stops.push(cur);
    return stops;
  }

  function parseExportDate(paragraphs) {
    for (const p of paragraphs) {
      const m = p.match(/匯出日期：(\d{4})-(\d{2})-(\d{2})/);
      if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    }
    return null;
  }

  async function parseRouteDocx(file) {
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const xmlStr = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xmlStr, "application/xml");
    const paragraphs = descendantsNS(doc.documentElement, W_NS, "p").map(paragraphText);
    return { stops: parseRouteParagraphs(paragraphs), exportDate: parseExportDate(paragraphs) };
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
  async function buildReport({ reportFile, routeFile, photoFiles, district, dateOverride, dryRun }) {
    const { stops, exportDate } = await parseRouteDocx(routeFile);
    const fallbackDate = dateOverride || (exportDate ? `${exportDate.m}/${exportDate.d}` : null);

    const photoByStem = {};
    for (const f of photoFiles) {
      const stem = f.name.replace(/\.(jpe?g)$/i, "");
      if (CASE_ID_RE.test(stem)) photoByStem[stem] = f;
    }

    const matched = stops.filter((s) => photoByStem[s.id]);
    if (matched.length !== 5) {
      throw new ReportError(
        `照片資料夾裡符合路線案件編號的照片有 ${matched.length} 張，需要剛好 5 張才能繼續。` +
        (matched.length ? `目前符合的案件：${matched.map((s) => s.id).join(", ")}` : "")
      );
    }
    const recheck = matched.filter((s) => s.isLastMonth);
    const fresh = matched.filter((s) => !s.isLastMonth);
    if (recheck.length !== 1 || fresh.length !== 4) {
      throw new ReportError(`需要剛好 1 筆「上月抽查」＋4 筆「本月」，目前是 ${recheck.length} 筆上月 + ${fresh.length} 筆本月。`);
    }
    fresh.sort((a, b) => a.order - b.order);
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
      const { text: addr, lowConfidence } = simplifyAddress(s.addressRaw);
      const stage = s.isLastMonth ? "複查" : "抽查";
      const closing = s.isLastMonth ? "，權責機關確已完成案件處理及抽查作業。" : "，權責機關確依限完成案件處理作業。";
      const contentLine2 = `經實地查證，${buildVerificationPhrase(s.category)}` + closing;
      const dateMd = rowDates[s.id];

      preview.push({
        row: i + 1, id: s.id, stage, category: s.category, agency: s.agency,
        date: dateMd, address: addr, addressLowConfidence: lowConfidence,
        addressRaw: s.addressRaw, photoName: photoByStem[s.id].name,
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

  window.AuditReportModule = { buildReport, simplifyAddress, parseRouteDocx, photoExifDate, ReportError };

  // ---------- UI ----------
  const KNOWN_DISTRICTS = ["中正", "萬華", "松山", "信義", "大安", "文山", "中山", "大同", "士林", "北投", "內湖", "南港"];

  function guessDistrict(filename) {
    return KNOWN_DISTRICTS.find((d) => filename.includes(d)) || "";
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
        <td>${escapeHtmlLocal(r.photoName)}</td>
      </tr>`).join("");
    return `<table class="case-table">
      <thead><tr><th>列</th><th>階段</th><th>案號</th><th>日期</th><th>項目</th><th>機關</th><th>簡化地址</th><th>照片</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  function escapeHtmlLocal(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  let lastBuildResult = null;

  async function runBuild(dryRun) {
    const statusEl = document.getElementById("reportStatus");
    const previewEl = document.getElementById("reportPreview");
    const reportInput = document.getElementById("reportFileInput");
    const routeInput = document.getElementById("reportRouteInput");
    const photosInput = document.getElementById("reportPhotosInput");
    const districtSel = document.getElementById("reportDistrict");
    const dateInput = document.getElementById("reportDateOverride");

    if (!reportInput.files[0]) { alert("請先選擇月報範本 docx。"); return; }
    if (!routeInput.files[0]) { alert("請先選擇查訪路線 docx。"); return; }
    const photoFiles = Array.from(photosInput.files || []).filter((f) => /\.jpe?g$/i.test(f.name));
    if (!photoFiles.length) { alert("請先選擇照片資料夾（裡面要有以案件編號命名的 jpg 照片）。"); return; }
    if (!districtSel.value) { alert("請選擇行政區。"); return; }

    statusEl.textContent = "處理中…";
    previewEl.innerHTML = "";
    try {
      const result = await window.AuditReportModule.buildReport({
        reportFile: reportInput.files[0],
        routeFile: routeInput.files[0],
        photoFiles,
        district: districtSel.value,
        dateOverride: dateInput.value.trim() || null,
        dryRun,
      });
      lastBuildResult = dryRun ? null : result;
      let warn = "";
      if (result.distinctDates.length > 1) {
        warn = `<div class="hint" style="color:#c9821a;margin-top:6px">⚠ 這 5 張照片拍攝日期不只一天：${result.distinctDates.join("、")}，已各自照實際拍攝日期填入，請確認這是預期中的狀況。</div>`;
      }
      previewEl.innerHTML = renderPreviewTable(result.preview) + warn;
      statusEl.textContent = dryRun ? "預覽完成（尚未產生檔案，確認沒問題後按下方「產生報告並下載」）" : "已產生報告，準備下載…";
      if (!dryRun && result.blob) downloadResult(result.blob, reportInput.files[0].name, districtSel.value);
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

  function downloadResult(blob, originalName, district) {
    const now = new Date();
    const ts = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;
    const base = originalName.replace(/\.docx$/i, "");
    const filename = `${base}_已更新_${district}_${ts}.docx`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  document.addEventListener("DOMContentLoaded", () => {
    const routeInput = document.getElementById("reportRouteInput");
    const districtSel = document.getElementById("reportDistrict");
    if (!routeInput || !districtSel) return; // 這個頁籤還沒建立在畫面上（理論上不會發生，防禦性檢查）

    KNOWN_DISTRICTS.forEach((d) => {
      const opt = document.createElement("option");
      opt.value = d; opt.textContent = d;
      districtSel.appendChild(opt);
    });

    // 「選擇...」按鈕轉發點擊到隱藏的 <input type=file>（跟資料匯入頁籤的既有按鈕/input 配對方式一樣），
    // 選好之後在旁邊的 hint 顯示檔名/張數，讓使用者確認選對了
    function wireFileButton(btnId, inputId, statusId, describe) {
      const btn = document.getElementById(btnId);
      const input = document.getElementById(inputId);
      const status = document.getElementById(statusId);
      btn.addEventListener("click", () => input.click());
      input.addEventListener("change", () => { status.textContent = describe(input.files); });
    }
    wireFileButton("btnChooseReportFile", "reportFileInput", "reportFileStatus",
      (files) => (files[0] ? `已選擇：${files[0].name}` : ""));
    wireFileButton("btnChooseReportRoute", "reportRouteInput", "reportRouteStatus",
      (files) => (files[0] ? `已選擇：${files[0].name}` : ""));
    wireFileButton("btnChooseReportPhotos", "reportPhotosInput", "reportPhotosStatus",
      (files) => {
        const jpgs = Array.from(files).filter((f) => /\.jpe?g$/i.test(f.name));
        return jpgs.length ? `已選擇資料夾，找到 ${jpgs.length} 張 jpg 照片` : "資料夾裡沒有找到 jpg 照片";
      });

    routeInput.addEventListener("change", () => {
      const f = routeInput.files[0];
      if (f) {
        const guess = guessDistrict(f.name);
        if (guess) districtSel.value = guess;
      }
    });

    document.getElementById("btnReportPreview").addEventListener("click", () => runBuild(true));
    document.getElementById("btnReportGenerate").addEventListener("click", () => runBuild(false));
  });
})();
