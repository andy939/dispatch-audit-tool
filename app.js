/* 派工案件抽查規劃工具 - 前端邏輯 (無需伺服器，純瀏覽器端執行) */
(function () {
  "use strict";

  const LS_CASES = "dispatch_cases_v1";
  const LS_GEOCACHE = "dispatch_geocache_v1";
  const LS_FOCUS = "dispatch_focus_v1"; // 雙向「跳到地圖/跳回清單」訊號，主視窗與地圖視窗共用同一個 key
  const LS_FOCUS_ORIGIN_TAB = "dispatch_focus_origin_tab_v1"; // 記錄是從左側哪個頁籤跳去地圖的，地圖那邊點別的案件跳回來時要回到同一個頁籤
  const LS_MARKED = "dispatch_marked_v1"; // 標記要排路線的案件
  const LS_REJECTED = "dispatch_rejected_v1"; // 標記「暫時否決、不列入抽查考量」的案件；只影響顯示樣式（見 toggleRejected）
  const LS_ROUTE = "dispatch_route_v1"; // 規劃好的路線，傳給地圖視窗畫線
  const LS_ROUTE_START = "dispatch_route_start_v1"; // 選定的起點捷運站名稱
  const LS_ROUTE_LOCK_END = "dispatch_route_lock_end_v1"; // 鎖定為「最後一站」的案件 id，規劃路線時固定排最後，其餘案件仍自動重新排序
  const LS_SETTINGS = "dispatch_settings_v1"; // 地理編碼設定 (Google API 金鑰、服務選擇)，主視窗與地圖視窗共用
  const LS_PAGE_SIZE = "dispatch_page_size_v1"; // 案件清單每頁筆數
  const LS_EXCLUDE_ON_IMPORT = "dispatch_exclude_old_garbage_on_import_v1"; // 匯入時是否自動排除3天前的垃圾/散落物案件（checkbox 記憶）
  const LS_LASTMONTH_FULL = "dispatch_lastmonth_full_v1"; // 「上月抽查管理」匯入的上月完整案件清單（原始資料，跟 state.cases 同 shape）
  const LS_LASTMONTH_AUDIT = "dispatch_lastmonth_audit_v1"; // 「上月抽查管理」匯入的機關抽查清單（原始資料，只有抽查專屬欄位）
  const LS_LASTMONTH_CASES = "dispatch_lastmonth_view_v1"; // getLastMonthCases() 算出來的合併結果，單向廣播給地圖視窗用，app.js 自己不靠這個 key 讀資料
  const LS_CATEGORY_FILTER = "dispatch_category_filter_v1"; // 「依案件項目」排除設定（state.filter.category），單向廣播給地圖視窗，讓地圖只畫出沒被排除的案件

  const DEFAULT_PAGE_SIZE = 10;

  // 跟 map.js 用同一份色盤（顏色/順序都要一樣），這樣同一個案件項目在主視窗徽章跟地圖點位
  // 才會是同一個顏色。使用者要求參考 Google 地圖的配色風格；Google 地圖本身的圖標配色沒有公開
  // 對外的固定色碼表，改採 Google 自家 Material Design 色盤（Google 產品一致採用的標準色系，
  // 同樣是這種飽和、乾淨的視覺風格），取 12 色的 500 色階，避開灰/黑（原因同上，怕跟捷運站圖例的
  // 灰色點混在一起）
  const COLORS = ["#F44336", "#E91E63", "#9C27B0", "#673AB7", "#3F51B5", "#2196F3",
    "#009688", "#4CAF50", "#FFC107", "#FF9800", "#FF5722", "#795548"];

  // 這兩項是抽查流程裡最後才查驗的項目，使用者要求固定用灰階跟其他項目的高飽和色區分開來，
  // 不要跟著雜湊自動配色。跟 map.js 要用同一份，兩邊必須完全一樣。
  const FIXED_COLORS = { "鄰里無主垃圾清運": "#cccccc", "道路散落物或油漬處理": "#4a4a4a" };

  // 查驗共識：這兩項案件項目只查 3 天內成案的，3 天前的直接不列入考量（不是否決標記、是真的從
  // 資料裡移除）。使用者一開始要用 🚫 否決標記處理，但案件量大時「清單裡還留著、只是變淡」一樣要
  // 捲過去找，改成真的刪除。原本還有一個「復原」按鈕可以把剛移除的一批找回來，使用者確認這兩項
  // 一旦超過 3 天就真的沒有查驗意義了，不會想復原，這個功能已經拿掉。
  const AUTO_EXCLUDE_CATEGORIES = ["鄰里無主垃圾清運", "道路散落物或油漬處理"];
  const AUTO_EXCLUDE_DAYS = 3;

  // 案件項目/承辦機關的篩選是多選勾選框，預設全部勾選（= 全部顯示）。
  // 為了讓「預設全部勾選」不用真的在畫面初始化時把所有值都塞進 Set，這裡改成反過來存「取消勾選=排除」
  // 的值：Set 是空的代表沒有排除任何東西（=全部顯示，checkbox 畫面上就會全部打勾）；
  // 取消勾選某個值就把它加進這個 Set，之後畫面上不會顯示、案件清單也會濾掉。
  // 不同欄位之間是 AND（例如取消勾了兩個案件項目 + 一個承辦機關 = 排除這兩個項目 且 排除這個機關）。
  // 原本還有 district（依行政區）跟 status（全部狀態）兩個篩選，使用者確認這個工具每次匯入的
  // xlsx 本來就是單一行政區的匯出檔、案件狀態幾乎全部是「已結案」，這兩個篩選從來沒有實際用途，
  // 拿掉了對應的 UI 跟這裡的欄位
  function emptyFilter() {
    return { category: new Set(), agency: new Set(), search: "" };
  }

  let state = {
    cases: [],
    marked: new Set(),
    rejected: new Set(),
    filter: emptyFilter(),
    sort: { key: "createdAt", dir: "desc" },
    geocache: {},
    page: 1,
    pageSize: DEFAULT_PAGE_SIZE,
    activeTab: "list",
    lastMonthFullList: [], // 「上月抽查管理」匯入的上月完整案件清單（原始資料）
    lastMonthAuditList: [], // 「上月抽查管理」匯入的機關抽查清單（原始資料）
    lastMonthSearch: "", // 上月抽查清單的搜尋關鍵字，跟 state.filter.search 一樣只影響「這裡顯示什麼」，不影響 getLastMonthCases() 本身（附近案件/依捷運站分群/路線規劃都不受這個搜尋影響）
    routeLockEndId: null, // 鎖定為「最後一站」的案件 id；null 代表沒有鎖定，維持原本自動排序不受影響
  };

  let mapWindowRef = null;
  let lastRoute = null; // { startStation, order } - 最近一次算好的路線，供匯出 CSV 用
  // 使用者自己在地圖視窗截圖存檔後選入的圖片，匯出 Word 時貼在文件最上方；純瀏覽器端沒辦法
  // 自動擷取地圖畫面（底圖圖磚來自 OSM/Google，兩邊都不允許用 canvas 讀取像素做截圖），
  // 所以改成讓使用者自己截圖、手動選檔這條路
  let routeMapImageData = null; // { bytes: Uint8Array, type: "png"|"jpg", width, height }

  // ---------- 工具函式 ----------
  // 案件項目名稱有些很長（例如「交通號誌電纜線垂落及設施損壞」），徽章換行會擠壓其他欄位，
  // 所以畫面上一律截到最多 10 個字＋「…」，完整名稱放 title 讓滑鼠移過去看
  function truncateCategory(name) {
    const s = name || "";
    return s.length > 10 ? s.slice(0, 10) + "…" : s;
  }

  // 依捷運站分群／附近案件／已標記案件這幾個地方現在會混著顯示本月案件跟上月抽查案件（getLastMonthCases()
  // 合併出來的案件 id 一律「抽」開頭），畫面上顯示的案件編號改用 realId（看不到內部前綴），
  // 所以要另外補一個小標籤讓使用者看得出「這筆是上月抽查的」，不然兩種案件混在一起會分不出來
  function lastMonthTagHtml(c) {
    return c.id.indexOf("抽") === 0 ? `<span class="lastmonth-tag" title="上月抽查案件">📋抽查</span> ` : "";
  }

  // 「案件內容」欄位的共用內容：案件清單／依捷運站分群／上月抽查管理三個表格都是同一套「短預覽 +
  // 展開內容」邏輯，抽成一個函式集中維護，改一次三個地方都同步，不用擔心哪裡漏改。
  // 「案件處理」（機關實際處理過程的記錄，有時候比案件內容本身更有參考價值）併進展開內容裡面顯示，
  // 不另開一欄——表格本來就 8 欄很緊繃，案件處理常常跟案件內容一樣長，收合起來單獨佔一欄也沒意義。
  // extraHtml 給「上月抽查管理」表格附加抽查資訊用（那段本來就設計成一律顯示、不跟著收合）
  function contentCellHtml(c, extraHtml) {
    const content = escapeHtml(c.content);
    const short = content.length > 40 ? content.slice(0, 40) + "…" : content;
    const handlingHtml = c.handling ? `<div class="hint" style="margin-top:4px"><b>案件處理：</b>${escapeHtml(c.handling)}</div>` : "";
    const needsToggle = content.length > 40 || !!c.handling;
    return `
      <div class="short">${short}</div>
      ${needsToggle ? `<span class="toggle">展開內容</span><div class="full">${content}${handlingHtml}</div>` : ""}
      ${extraHtml || ""}
    `;
  }

  function colorFor(name) {
    if (FIXED_COLORS[name]) return FIXED_COLORS[name];
    let h = 0;
    for (let i = 0; i < (name || "").length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
    return COLORS[h % COLORS.length];
  }

  function escapeHtml(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  // 「成案時間」是派工系統匯出的民國年格式，例如「115/08/18 13:12」或「115/08/18 13:12:09」，
  // 沒有時間部分時也要能解析（當成當天 00:00）。解析不出來就回傳 null，呼叫端要自己決定怎麼跳過，
  // 不要用「解析失敗當作最舊/最新」這種猜測，避免誤判案件的新舊
  function parseRocDateTime(s) {
    const m = String(s || "").trim().match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (!m) return null;
    const [, rocYear, month, day, hour, minute, second] = m;
    const d = new Date(
      Number(rocYear) + 1911, Number(month) - 1, Number(day),
      Number(hour || 0), Number(minute || 0), Number(second || 0)
    );
    return isNaN(d.getTime()) ? null : d;
  }

  // 全形數字/英文/符號轉半形，需與地圖視窗 (map.js) 的地理編碼快取 key 規則一致
  function toHalfWidth(s) {
    return (s || "")
      .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/　/g, " ");
  }

  function cacheKeyFor(c) {
    return (c.district || "") + "|" + toHalfWidth(c.address).trim();
  }

  function haversine(lat1, lon1, lat2, lon2) {
    const R = 6371000;
    const toRad = (d) => (d * Math.PI) / 180;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  }

  function nearestStation(lat, lng) {
    let best = null, bestDist = Infinity;
    for (const st of (window.MRT_STATIONS || [])) {
      const d = haversine(lat, lng, st.lat, st.lng);
      if (d < bestDist) { bestDist = d; best = st; }
    }
    return best ? { station: best, distance: bestDist } : null;
  }

  // ---------- 路線規劃：點到點盡量不要回頭 (小規模最近鄰 + 2-opt 改善) ----------
  function routeDistance(points, order) {
    let d = 0;
    for (let i = 0; i < order.length - 1; i++) {
      d += haversine(points[order[i]].lat, points[order[i]].lng, points[order[i + 1]].lat, points[order[i + 1]].lng);
    }
    return d;
  }

  // lastIdx 有給值時，那個點會被排到最後才拜訪（除非它是唯一剩下沒排的點，那時不得不選它）——
  // 供「鎖定最後一站」使用，讓最近鄰居法找路徑時不會提早經過那個點
  function nearestNeighborOrder(points, startIdx, lastIdx) {
    const n = points.length;
    const visited = new Array(n).fill(false);
    const order = [startIdx];
    visited[startIdx] = true;
    for (let step = 1; step < n; step++) {
      const last = points[order[order.length - 1]];
      const remaining = n - step;
      let best = -1, bestDist = Infinity;
      for (let i = 0; i < n; i++) {
        if (visited[i]) continue;
        if (lastIdx != null && i === lastIdx && remaining > 1) continue;
        const d = haversine(last.lat, last.lng, points[i].lat, points[i].lng);
        if (d < bestDist) { bestDist = d; best = i; }
      }
      order.push(best);
      visited[best] = true;
    }
    return order;
  }

  // 2-opt：反覆嘗試把路線中兩段交換方向，只要能縮短總距離就採用，直到沒有更好的為止 (消除「繞回頭」的交叉路徑)。
  // keepEndFixed 為 true 時，交換範圍不含最後一個點——供「鎖定最後一站」使用，2-opt 改善其餘路段時不會把鎖定的終點換到別的位置
  function twoOptImprove(points, order, keepEndFixed) {
    let best = order.slice();
    let bestDist = routeDistance(points, best);
    const kLimit = keepEndFixed ? best.length - 1 : best.length;
    let improved = true;
    while (improved) {
      improved = false;
      for (let i = 1; i < best.length - 1; i++) {
        for (let k = i + 1; k < kLimit; k++) {
          const candidate = best.slice(0, i).concat(best.slice(i, k + 1).reverse(), best.slice(k + 1));
          const d = routeDistance(points, candidate);
          if (d < bestDist - 1e-6) { best = candidate; bestDist = d; improved = true; }
        }
      }
    }
    return best;
  }

  function permute(arr, l, onFull) {
    if (l === arr.length - 1) { onFull(arr); return; }
    for (let i = l; i < arr.length; i++) {
      [arr[l], arr[i]] = [arr[i], arr[l]];
      permute(arr, l + 1, onFull);
      [arr[l], arr[i]] = [arr[i], arr[l]];
    }
  }

  // points: [{id, lat, lng}, ...]；回傳造訪順序的 index 陣列
  // fixedStartIdx 有給值時，路線一定從那個點出發 (例如指定的起點捷運站)，不會被排到別的位置
  // fixedEndIdx 有給值時，路線一定以那個點結束 (例如使用者鎖定的「最後一站」)，不會被排到別的位置；
  // 兩者可以同時給值 (起點+終點都固定，中間其他案件仍自動排最短路徑)
  function planRouteOrder(points, fixedStartIdx, fixedEndIdx) {
    const n = points.length;
    if (n <= 1) return points.map((_, i) => i);
    const hasStart = fixedStartIdx != null;
    const hasEnd = fixedEndIdx != null && fixedEndIdx !== fixedStartIdx;

    if (n <= 8) {
      // 案件不多時直接窮舉所有排列，保證最短；固定起點/終點的話，只窮舉排列中間那些自由的點
      const fixedSet = new Set([hasStart ? fixedStartIdx : null, hasEnd ? fixedEndIdx : null].filter((x) => x != null));
      const freeIdxs = points.map((_, i) => i).filter((i) => !fixedSet.has(i));
      if (!freeIdxs.length) return (hasStart ? [fixedStartIdx] : []).concat(hasEnd ? [fixedEndIdx] : []);
      let bestOrder = null, bestDist = Infinity;
      permute(freeIdxs.slice(), 0, (arr) => {
        const order = (hasStart ? [fixedStartIdx] : []).concat(arr).concat(hasEnd ? [fixedEndIdx] : []);
        const d = routeDistance(points, order);
        if (d < bestDist) { bestDist = d; bestOrder = order; }
      });
      return bestOrder;
    }

    if (hasEnd) {
      // 案件較多時：終點固定，起點固定的話只試那一個，否則每個其他點都試著當起點做最近鄰（終點保留到最後才拜訪），再用 2-opt 改善（不動終點），取最好的一組
      const starts = hasStart ? [fixedStartIdx] : points.map((_, i) => i).filter((i) => i !== fixedEndIdx);
      let bestOrder = null, bestDist = Infinity;
      starts.forEach((s) => {
        const order = twoOptImprove(points, nearestNeighborOrder(points, s, fixedEndIdx), true);
        const d = routeDistance(points, order);
        if (d < bestDist) { bestDist = d; bestOrder = order; }
      });
      return bestOrder;
    }

    if (hasStart) return twoOptImprove(points, nearestNeighborOrder(points, fixedStartIdx));

    // 案件較多、起點終點都沒鎖定：每個點都試著當起點做最近鄰，再用 2-opt 改善，取最好的一組
    let bestOrder = null, bestDist = Infinity;
    for (let s = 0; s < n; s++) {
      const order = twoOptImprove(points, nearestNeighborOrder(points, s));
      const d = routeDistance(points, order);
      if (d < bestDist) { bestDist = d; bestOrder = order; }
    }
    return bestOrder;
  }

  // ---------- 與地圖視窗雙向定位 (透過 localStorage 的 LS_FOCUS 這個 key 互相通知) ----------
  function openMapWindow() {
    if (mapWindowRef && !mapWindowRef.closed) { mapWindowRef.focus(); return mapWindowRef; }
    mapWindowRef = window.open("map.html", "dispatchMap", "width=1100,height=880");
    return mapWindowRef;
  }

  function focusCaseOnMap(id) {
    const c = state.cases.find((x) => x.id === id) || getLastMonthCases().find((x) => x.id === id);
    if (!c) return;
    const loc = state.geocache[cacheKeyFor(c)];
    if (!loc) {
      alert("這筆案件尚未在地圖上定位，請先開啟地圖視窗並按「定位案件座標」。");
      return;
    }
    // 記住是從哪個頁籤點過去地圖的，之後在地圖上點別的案件跳回主視窗時才知道要回到同一個頁籤
    localStorage.setItem(LS_FOCUS_ORIGIN_TAB, state.activeTab);
    localStorage.setItem(LS_FOCUS, JSON.stringify({ id, ts: Date.now() }));
    renderNearbyPanel(id); // 點案件當下就馬上同步周邊清單，不用等使用者再跑去地圖上點一次才看到
    openMapWindow();
  }

  // ---------- 左側頁籤切換 ----------
  function switchTab(tab) {
    state.activeTab = tab;
    document.querySelectorAll(".side-nav-item").forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tab));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.toggle("active", panel.dataset.tabPanel === tab));
  }

  // 在地圖上點了某案件標記時：切到「依捷運站分群」頁籤，展開該案件所在的站群、捲動並閃爍標示
  // （地圖上的標記一定是已經定位過的案件，所以一定會落在某個站群裡）
  function revealCaseInStationGroup(id) {
    switchTab("station");
    const row = document.querySelector(`#stationGroupPanel .case-row[data-case-id="${CSS.escape(id)}"]`);
    if (!row) return;
    const group = row.closest("details.station-group");
    if (group) group.open = true;
    document.querySelectorAll("#stationGroupPanel .flash-highlight").forEach((r) => r.classList.remove("flash-highlight"));
    row.scrollIntoView({ behavior: "smooth", block: "center" });
    row.classList.add("flash-highlight");
  }

  // 切到「本月案件管理」頁籤，找到該筆案件、跳到所在頁、捲動並閃爍標示
  function revealCaseInList(id) {
    switchTab("list");
    let tr = document.querySelector(`#caseTableBody tr[data-id="${CSS.escape(id)}"]`);
    if (!tr) {
      state.filter = emptyFilter();
      document.getElementById("searchBox").value = "";
      const fullList = applyFilters();
      const idx = fullList.findIndex((c) => c.id === id);
      state.page = idx >= 0 ? Math.floor(idx / state.pageSize) + 1 : 1;
      renderAll();
      tr = document.querySelector(`#caseTableBody tr[data-id="${CSS.escape(id)}"]`);
    }
    if (!tr) return;
    document.querySelectorAll("#caseTableBody tr.flash-highlight").forEach((r) => r.classList.remove("flash-highlight"));
    tr.scrollIntoView({ behavior: "smooth", block: "center" });
    tr.classList.add("flash-highlight");
  }

  // 切到「路線規劃」頁籤，找到該筆案件在路線裡的那一列、捲動並閃爍標示；
  // 如果還沒按過「規劃路線」（routePanel 裡沒有排好的那一列），退而求其次找「已標記案件」清單裡的列
  function revealCaseInRoute(id) {
    switchTab("route");
    let row = document.querySelector(`#routePanel .route-row[data-case-id="${CSS.escape(id)}"]`);
    if (!row) row = document.querySelector(`#markedListPanel tr[data-id="${CSS.escape(id)}"]`);
    if (!row) return;
    document.querySelectorAll("#routePanel .flash-highlight, #markedListPanel .flash-highlight").forEach((r) => r.classList.remove("flash-highlight"));
    row.scrollIntoView({ behavior: "smooth", block: "center" });
    row.classList.add("flash-highlight");
  }

  // 切到「上月抽查管理」頁籤，找到該筆案件、捲動並閃爍標示
  function revealCaseInLastMonth(id) {
    switchTab("lastmonth");
    const tr = document.querySelector(`#lastMonthTableBody tr[data-id="${CSS.escape(id)}"]`);
    if (!tr) return;
    document.querySelectorAll("#lastMonthTableBody tr.flash-highlight").forEach((r) => r.classList.remove("flash-highlight"));
    tr.scrollIntoView({ behavior: "smooth", block: "center" });
    tr.classList.add("flash-highlight");
  }

  // 地圖那邊點了案件標記跳回主視窗時，依「原本是從哪個頁籤點過去地圖的」決定要跳回哪個頁籤——
  // 不管去程是從哪裡點過去，回程都應該回到同一個地方，而不是每次都固定跳去某個頁籤
  function revealCaseByOrigin(id) {
    const originTab = localStorage.getItem(LS_FOCUS_ORIGIN_TAB);
    if (originTab === "list") revealCaseInList(id);
    else if (originTab === "route") revealCaseInRoute(id);
    else if (originTab === "lastmonth") revealCaseInLastMonth(id);
    else revealCaseInStationGroup(id); // 預設 / 從「依捷運站分群」點過去的情況
  }

  // ---------- 資料載入 / 儲存 ----------
  // 修正舊資料，因為匯入預設是「合併現有資料」，已經存在 localStorage 裡的舊案件不會被
  // 後來版本改進過的解析邏輯自動清乾淨，所以每次載入時強制清一次：
  // 1. 案件編號欄位如果原始 xlsx 裡是多行內容，早期版本沒有把 id 修乾淨，尾端可能還留著
  //    換行字元（例如 "2026081710254\n"），拿 dataset.id 去比對會永遠比不到。
  // 2. 地址欄位如果有全形數字/符號，地理編碼常常會定位到奇怪的地方（跟以前手動用 Excel
  //    ASC() 函數轉半形再匯入 Google 我的地圖是同一個道理），現在改成匯入時自動轉半形，
  //    但這只對「以後新匯入」的案件有效，已經存在的舊案件也要在這裡一併轉乾淨。
  function migrateCases(cases) {
    let changed = false;
    cases.forEach((c) => {
      const cleanId = String(c.id || "").trim();
      const cleanIdRaw = String(c.idRaw || "").trim();
      const cleanAddress = toHalfWidth(c.address || "").trim();
      if (cleanId !== c.id) { c.id = cleanId; changed = true; }
      if (cleanIdRaw !== c.idRaw) { c.idRaw = cleanIdRaw; changed = true; }
      if (cleanAddress !== c.address) { c.address = cleanAddress; changed = true; }
    });
    return changed;
  }

  function loadState() {
    let cases = null;
    try { cases = JSON.parse(localStorage.getItem(LS_CASES) || "null"); } catch (e) {}
    state.cases = cases || [];
    if (migrateCases(state.cases)) persistCases();

    try { state.geocache = JSON.parse(localStorage.getItem(LS_GEOCACHE) || "{}"); } catch (e) { state.geocache = {}; }
    try { state.marked = new Set(JSON.parse(localStorage.getItem(LS_MARKED) || "[]")); } catch (e) { state.marked = new Set(); }
    try { state.rejected = new Set(JSON.parse(localStorage.getItem(LS_REJECTED) || "[]")); } catch (e) { state.rejected = new Set(); }
    const savedPageSize = parseInt(localStorage.getItem(LS_PAGE_SIZE), 10);
    state.pageSize = [10, 20, 50].includes(savedPageSize) ? savedPageSize : DEFAULT_PAGE_SIZE;
    try { state.lastMonthFullList = JSON.parse(localStorage.getItem(LS_LASTMONTH_FULL) || "[]"); } catch (e) { state.lastMonthFullList = []; }
    try { state.lastMonthAuditList = JSON.parse(localStorage.getItem(LS_LASTMONTH_AUDIT) || "[]"); } catch (e) { state.lastMonthAuditList = []; }
    state.routeLockEndId = localStorage.getItem(LS_ROUTE_LOCK_END) || null;
    // 現在地圖視窗的圖例圓點也能改這個排除設定（見 toggleCategoryFilterFromMap()），重新整理主視窗時
    // 也要照這個值還原，不然剛好在地圖那邊點過篩選後，回主視窗重新整理會看到兩邊的篩選狀態對不起來
    try { state.filter.category = new Set(JSON.parse(localStorage.getItem(LS_CATEGORY_FILTER) || "[]")); } catch (e) { state.filter.category = new Set(); }
  }

  function persistCases() {
    localStorage.setItem(LS_CASES, JSON.stringify(state.cases));
  }
  function persistLastMonthFullList() {
    localStorage.setItem(LS_LASTMONTH_FULL, JSON.stringify(state.lastMonthFullList));
  }
  function persistLastMonthAuditList() {
    localStorage.setItem(LS_LASTMONTH_AUDIT, JSON.stringify(state.lastMonthAuditList));
  }
  // 單向廣播給地圖視窗用（地圖只讀這個 key，不參與合併邏輯），app.js 自己一律用 getLastMonthCases() 現算
  function broadcastLastMonthView() {
    localStorage.setItem(LS_LASTMONTH_CASES, JSON.stringify(getLastMonthCases()));
  }
  function persistGeocache() {
    localStorage.setItem(LS_GEOCACHE, JSON.stringify(state.geocache));
  }
  function persistMarked() {
    localStorage.setItem(LS_MARKED, JSON.stringify(Array.from(state.marked)));
  }
  function persistRejected() {
    localStorage.setItem(LS_REJECTED, JSON.stringify(Array.from(state.rejected)));
  }
  function persistRouteLockEnd() {
    if (state.routeLockEndId) localStorage.setItem(LS_ROUTE_LOCK_END, state.routeLockEndId);
    else localStorage.removeItem(LS_ROUTE_LOCK_END);
  }
  // 「已標記案件」清單上按 📌 選一筆案件鎖定為最後一站；再按一次同一筆或案件被取消標記時解除。
  // 鎖定只影響「規劃路線」時終點固定在哪個案件，其餘案件仍會照常自動排出最短路徑——
  // 使用者確認這是刻意的：想先手動拖曳試一下順序，不確定是不是最佳解時，可以鎖住想放最後的那一站，
  // 再放心按「規劃路線」讓其他站重新最佳化，不會把這一站的位置洗掉
  function toggleRouteLockEnd(id) {
    state.routeLockEndId = state.routeLockEndId === id ? null : id;
    persistRouteLockEnd();
    renderMarkedList();
  }
  function persistRouteStart(name) {
    localStorage.setItem(LS_ROUTE_START, name || "");
  }

  // ---------- xlsx 解析 (與 build_data.py 對應欄位) ----------
  function cleanStr(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function parseWorkbookFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: false });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
          // 系統原始匯出檔前幾列通常是報表標題/列印資訊，真正的欄位標題列不一定在第 1 列，
          // 因此掃描找出含有「案件編號」的那一列當作標題列，其餘略過。
          const headerRowIndex = rows.findIndex((r) => r && r.some((cell) => String(cell).trim() === "案件編號"));
          if (headerRowIndex === -1) {
            throw new Error("找不到「案件編號」欄位標題，請確認這是派工系統匯出的案件清單");
          }
          const header = rows[headerRowIndex];
          const idx = {};
          header.forEach((h, i) => { idx[String(h).trim()] = i; });
          const get = (r, key) => (idx[key] !== undefined ? cleanStr(r[idx[key]]) : "");
          const out = [];
          for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const r = rows[i];
            if (!r || !r.length) continue;
            const idRaw = get(r, "案件編號");
            if (!idRaw) continue;
            const addr2 = get(r, "發生地址2");
            const addr1 = get(r, "發生地址");
            out.push({
              id: idRaw.split("\n")[0].trim(),
              idRaw,
              status: get(r, "案件狀態"),
              category: get(r, "案件項目"),
              agency: get(r, "承辦機關"),
              district: get(r, "行政區"),
              address: toHalfWidth(addr2 || addr1).trim(), // 全形數字/符號轉半形（跟以前手動用 Excel ASC() 函數再匯入 Google 我的地圖的效果一樣），不然地理編碼常常會定位到奇怪的地方
              createdAt: get(r, "成案時間").replace(/\n/g, " "),
              dueAt: get(r, "限辦時間").replace(/\n/g, " "),
              closedAt: get(r, "結案時間").replace(/\n/g, " "),
              content: get(r, "案件內容"),
              handling: get(r, "案件處理"),
              level: get(r, "案件等級"),
              registrar: get(r, "立案人員"),
              satisfaction: get(r, "滿意度選項"),
              feedback: get(r, "民眾意見"),
              sourceFile: file.name,
            });
          }
          resolve(out);
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // 機關抽查清單是很單純的一份匯出（只有案件編號/項目/行政區 + 抽查專屬欄位，沒有地址/內容），
  // 標題列掃描邏輯跟 parseWorkbookFile() 一樣（一樣是找含「案件編號」的那一列），但欄位換成
  // 抽查清單獨有的那 7 個欄位，不能直接呼叫 parseWorkbookFile() 沿用（欄位對不上）
  function parseAuditWorkbookFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: false });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
          const headerRowIndex = rows.findIndex((r) => r && r.some((cell) => String(cell).trim() === "案件編號"));
          if (headerRowIndex === -1) {
            throw new Error("找不到「案件編號」欄位標題，請確認這是機關抽查清單的匯出檔");
          }
          const header = rows[headerRowIndex];
          const idx = {};
          header.forEach((h, i) => { idx[String(h).trim()] = i; });
          const get = (r, key) => (idx[key] !== undefined ? cleanStr(r[idx[key]]) : "");
          const out = [];
          for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const r = rows[i];
            if (!r || !r.length) continue;
            const idRaw = get(r, "案件編號");
            if (!idRaw) continue;
            out.push({
              id: idRaw.split("\n")[0].trim(),
              idRaw,
              category: get(r, "案件項目"),
              district: get(r, "行政區"),
              auditor: get(r, "抽查人員"),
              verifiedAt: get(r, "查證時間").replace(/\n/g, " "),
              signedAt: get(r, "抽查簽收時間").replace(/\n/g, " "),
              repliedAt: get(r, "抽查回覆時間").replace(/\n/g, " "),
              passedAt: get(r, "抽查通過時間").replace(/\n/g, " "),
              auditDueAt: get(r, "抽查限辦時間").replace(/\n/g, " "),
              auditNote: get(r, "抽查內容"),
              sourceFile: file.name,
            });
          }
          resolve(out);
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // 「上月抽查管理」的合併清單：以抽查清單（機關實際抽查過的案件）為準，其餘案件不列入；
  // 每筆用 id 去上月完整清單找明細（地址/內容/承辦機關等），找不到就 matched:false，畫面上會提示。
  // id 統一加上「抽」字前綴（realId 保留原始編號）：跟本月案件的 id 空間完全分開，
  // 避免萬一編號重複時 state.marked/state.rejected/geocache 互相污染，地圖/清單上也一眼看得出是上月資料
  function getLastMonthCases() {
    const fullById = new Map(state.lastMonthFullList.map((c) => [c.id, c]));
    return state.lastMonthAuditList.map((a) => {
      const full = fullById.get(a.id);
      return {
        id: "抽" + a.id,
        realId: a.id,
        idRaw: a.idRaw,
        category: a.category || (full && full.category) || "",
        district: a.district || (full && full.district) || "",
        agency: full ? full.agency : "",
        address: full ? full.address : "",
        content: full ? full.content : "",
        handling: full ? full.handling : "",
        createdAt: full ? full.createdAt : "",
        status: full ? full.status : "",
        matched: !!full,
        auditor: a.auditor, verifiedAt: a.verifiedAt, signedAt: a.signedAt,
        repliedAt: a.repliedAt, passedAt: a.passedAt, auditDueAt: a.auditDueAt, auditNote: a.auditNote,
      };
    });
  }

  async function handleFiles(fileList, mode) {
    const files = Array.from(fileList);
    if (!files.length) return;
    const excludeOldOnImport = document.getElementById("excludeOldOnImport").checked;
    const cutoff = autoExcludeCutoff();
    let merged = mode === "merge" ? state.cases.slice() : [];
    const seen = new Set(merged.map((c) => c.id));
    const excluded = [];
    for (const f of files) {
      try {
        const rows = await parseWorkbookFile(f);
        for (const c of rows) {
          if (seen.has(c.id)) continue;
          seen.add(c.id);
          if (excludeOldOnImport && isAutoExcludeCase(c, cutoff)) { excluded.push(c); continue; }
          merged.push(c);
        }
      } catch (err) {
        alert("解析檔案失敗：" + f.name + "\n" + err.message);
      }
    }
    state.cases = merged;
    persistCases();
    render();
    alert("匯入完成，目前共 " + state.cases.length + " 筆案件。" + (excluded.length ? `（已排除 ${excluded.length} 筆 ${AUTO_EXCLUDE_DAYS} 天前的垃圾/散落物案件）` : ""));
  }

  // 「上月抽查管理」的兩個匯入：跟主匯入不同，這裡是單一份「上月 snapshot」，
  // 每次選檔都是整批取代，不做合併模式（沒有「合併/取代上月清單」的意義，上月資料本來就是固定的）
  async function handleLastMonthFullFiles(fileList) {
    const files = Array.from(fileList);
    if (!files.length) return;
    const merged = [];
    const seen = new Set();
    for (const f of files) {
      try {
        const rows = await parseWorkbookFile(f);
        for (const c of rows) {
          if (seen.has(c.id)) continue;
          seen.add(c.id);
          merged.push(c);
        }
      } catch (err) {
        alert("解析檔案失敗：" + f.name + "\n" + err.message);
      }
    }
    state.lastMonthFullList = merged;
    persistLastMonthFullList();
    renderLastMonthPanel();
    alert("已匯入上月案件清單，共 " + merged.length + " 筆。");
  }

  async function handleLastMonthAuditFiles(fileList) {
    const files = Array.from(fileList);
    if (!files.length) return;
    const merged = [];
    const seen = new Set();
    for (const f of files) {
      try {
        const rows = await parseAuditWorkbookFile(f);
        for (const c of rows) {
          if (seen.has(c.id)) continue;
          seen.add(c.id);
          merged.push(c);
        }
      } catch (err) {
        alert("解析檔案失敗：" + f.name + "\n" + err.message);
      }
    }
    state.lastMonthAuditList = merged;
    persistLastMonthAuditList();
    renderLastMonthPanel();
    alert("已匯入抽查清單，共 " + merged.length + " 筆。");
  }

  // ---------- 統計 ----------
  function countBy(list, key) {
    const map = new Map();
    for (const c of list) {
      const k = c[key] || "（未填）";
      map.set(k, (map.get(k) || 0) + 1);
    }
    return Array.from(map.entries()).sort((a, b) => b[1] - a[1]);
  }

  function renderStatCards() {
    const cases = state.cases;
    const total = cases.length;
    const done = cases.filter((c) => (c.status || "").includes("結案")).length;
    const districts = new Set(cases.map((c) => c.district)).size;
    const categories = new Set(cases.map((c) => c.category)).size;
    const el = document.getElementById("statCards");
    el.innerHTML = [
      ["總案件數", total],
      ["已結案", done],
      ["未結案", total - done],
      ["行政區數", districts],
      ["案件項目種類", categories],
      ["已標記路線", state.marked.size],
    ].map(([label, num]) => `
      <div class="stat-card"><div class="num">${num}</div><div class="label">${label}</div></div>
    `).join("");

    const files = Array.from(new Set(cases.map((c) => c.sourceFile).filter(Boolean)));
    document.getElementById("dataSummaryLine").textContent =
      total ? `共 ${total} 筆案件，來源檔案：${files.join("、") || "內建範例"}` : "尚未載入資料";
  }

  function renderBreakdown(containerId, list, key, filterField) {
    const el = document.getElementById(containerId);
    const entries = countBy(list, key);
    const max = entries.length ? entries[0][1] : 1;
    const excluded = state.filter[filterField]; // Set 裡的值代表「不要顯示」，預設空集合 = 全部勾選/全部顯示
    // 使用者要求地圖也只顯示案件項目篩選後的結果（例如點色點只看某一類），這裡順便把目前的排除
    // 設定廣播給地圖視窗；每次重畫這個清單就重新廣播一次，不用額外去每個會改到 state.filter.category
    // 的地方另外加一行，任何變動最終都會經過這裡（renderAll() 一定會呼叫 renderBreakdown）
    if (filterField === "category") {
      localStorage.setItem(LS_CATEGORY_FILTER, JSON.stringify(Array.from(excluded)));
    }
    el.innerHTML = entries.map(([name, count]) => {
      const active = !excluded.has(name); // active = 有勾選 = 會顯示
      const pct = Math.round((count / max) * 100);
      const swatch = filterField === "category"
        ? `<span class="swatch swatch-clickable" title="點一下只顯示這個類別，再點一次恢復全部顯示" style="background:${colorFor(name)}"></span>`
        : "";
      return `
        <div class="bar-row ${active ? "active" : ""}" data-field="${filterField}" data-value="${escapeHtml(name)}">
          <div class="name"><input type="checkbox" class="bar-check" ${active ? "checked" : ""} tabindex="-1">${swatch}${escapeHtml(name)}</div>
          <div class="count">${count} 件</div>
          <div class="track"><div class="fill" style="width:${pct}%"></div></div>
        </div>`;
    }).join("") || `<div class="empty-state">尚無資料</div>`;

    const hintEl = document.getElementById(containerId.replace("breakdown", "hint"));
    if (hintEl) hintEl.textContent = entries.length ? `${entries.length} 種` : "";
    const detailsEl = document.getElementById(containerId.replace("breakdown", "details"));
    if (detailsEl && excluded.size) detailsEl.open = true;

    el.querySelectorAll(".bar-row").forEach((row) => {
      row.addEventListener("click", () => {
        const field = row.dataset.field;
        const value = row.dataset.value;
        const set = state.filter[field];
        if (set.has(value)) set.delete(value); else set.add(value); // 取消勾選 = 加入排除集合
        state.page = 1;
        renderAll();
      });
    });

    // 使用者要求比照「皮克敏名冊」那種篩選手感：點色點只顯示這一個類別（其他全部排除），
    // 再點一次（已經是「只剩這個」的狀態時）恢復全部顯示。直接用既有每列前面的色點當按鈕，
    // 不用另外多佔一排版面；跟色點所在的整列既有的「單項排除」點擊行為分開，要 stopPropagation
    // 不然點色點會同時觸發外層 .bar-row 的點擊，兩個行為互相打架
    el.querySelectorAll(".swatch-clickable").forEach((sw) => {
      sw.addEventListener("click", (e) => {
        e.stopPropagation();
        const row = sw.closest(".bar-row");
        const field = row.dataset.field;
        const value = row.dataset.value;
        const set = state.filter[field];
        const allNames = entries.map(([n]) => n);
        const isolated = allNames.length > 1 && !set.has(value) && allNames.every((n) => n === value || set.has(n));
        set.clear();
        if (!isolated) allNames.forEach((n) => { if (n !== value) set.add(n); });
        state.page = 1;
        renderAll();
      });
    });
  }

  // 案件項目/承辦機關的排除設定：案件清單表格跟依捷運站分群清單共用同一份 state.filter，
  // 排除某個類別後兩邊清單都要一致濾掉，不能一邊看得到一邊看不到。地圖跟路線規劃刻意不套用這個排除，
  // 維持「全部已匯入案件永遠同步」的既有設計（避免跟 ♡ 標記路線的機制搞混，見 README）。
  function applyCategoryFilters(list) {
    const { category, agency } = state.filter;
    if (category.size) list = list.filter((c) => !category.has(c.category || "（未填）"));
    if (agency.size) list = list.filter((c) => !agency.has(c.agency || "（未填）"));
    return list;
  }

  function applyFilters() {
    let list = applyCategoryFilters(state.cases);
    const { search } = state.filter;
    if (search) {
      const s = search.toLowerCase();
      list = list.filter((c) =>
        (c.address || "").toLowerCase().includes(s) ||
        (c.content || "").toLowerCase().includes(s) ||
        (c.id || "").toLowerCase().includes(s)
      );
    }
    const { key, dir } = state.sort;
    list = list.slice().sort((a, b) => {
      const av = a[key] || "", bv = b[key] || "";
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return list;
  }

  function renderTable() {
    const fullList = applyFilters();
    const tbody = document.getElementById("caseTableBody");
    const totalPages = Math.max(1, Math.ceil(fullList.length / state.pageSize));
    if (state.page > totalPages) state.page = totalPages;
    if (state.page < 1) state.page = 1;

    if (!fullList.length) {
      const msg = state.cases.length ? "沒有符合條件的案件" : "尚未匯入任何資料，請點右上角「匯入本月 xlsx 案件檔」開始";
      tbody.innerHTML = `<tr><td colspan="9"><div class="empty-state">${msg}</div></td></tr>`;
      renderPagination(0, 1);
      return;
    }
    const startIdx = (state.page - 1) * state.pageSize;
    const list = fullList.slice(startIdx, startIdx + state.pageSize);
    tbody.innerHTML = list.map((c) => {
      return `
        <tr data-id="${escapeHtml(c.id)}" class="${state.rejected.has(c.id) ? "row-rejected" : ""}">
          <td class="case-id-cell" title="${escapeHtml(c.idRaw)}">${escapeHtml(c.id)}</td>
          <td>${escapeHtml(c.district)}</td>
          <td><span class="badge" title="${escapeHtml(c.category)}" style="background:${colorFor(c.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(c.category))}</span></td>
          <td>${escapeHtml(c.agency)}</td>
          <td class="addr-cell-wrap" data-id="${escapeHtml(c.id)}" title="點這裡編輯地址"><div class="addr-cell-inner"><span class="addr-edit-btn">✏️</span><span class="addr-text">${escapeHtml(c.address)}</span></div></td>
          <td class="content-cell">${contentCellHtml(c)}</td>
          <td>${escapeHtml(c.createdAt)}</td>
          <td class="heart-td"><div class="heart-cell">
            <span class="heart-btn${state.marked.has(c.id) ? " marked" : ""}" data-case-id="${escapeHtml(c.id)}" title="標記要排路線">${state.marked.has(c.id) ? "❤️" : "🤍"}</span>
            <span class="reject-btn${state.rejected.has(c.id) ? " rejected" : ""}" data-case-id="${escapeHtml(c.id)}" title="${state.rejected.has(c.id) ? "取消否決（恢復正常顯示）" : "標記為暫時否決，不列入抽查考量（只影響顯示，不影響其他功能）"}">🚫</span>
          </div></td>
        </tr>`;
    }).join("");

    tbody.querySelectorAll(".toggle").forEach((t) => {
      t.addEventListener("click", (e) => {
        const cell = e.target.closest(".content-cell");
        cell.classList.toggle("expanded");
        e.target.textContent = cell.classList.contains("expanded") ? "收合內容" : "展開內容";
      });
    });

    tbody.querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", (e) => {
        if (e.target.closest("input, .toggle, .heart-btn, .reject-btn, .addr-edit-btn, .addr-save, .addr-cancel, .addr-cell-wrap")) return;
        tr.classList.remove("flash-highlight"); // 點了就當作已看到，手動取消反白
        focusCaseOnMap(tr.dataset.id);
      });
    });

    // 地址欄改成整格都能點進去編輯，不要求一定要點到那個小小的 ✏️ 圖示——
    // 地址字一長就會換行，鉛筆位置會跟著變，精準點到圖示很容易失手，看起來像沒反應。
    // 整格範圍大很多，點哪裡都會進入編輯，就不會有這個問題
    tbody.querySelectorAll(".addr-cell-wrap").forEach((td) => {
      td.addEventListener("click", (e) => {
        e.stopPropagation();
        if (td.querySelector(".addr-edit-input")) return;
        startEditAddress(td);
      });
    });

    tbody.querySelectorAll(".heart-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleMarked(btn.dataset.caseId);
      });
    });
    tbody.querySelectorAll(".reject-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleRejected(btn.dataset.caseId);
      });
    });

    renderPagination(fullList.length, totalPages);
  }

  // 「上月抽查管理」頁籤：合併清單通常一個行政區 40-50 筆、頂多百來筆，不特別做分頁。
  // 欄位格式跟案件清單一致，但地址/內容是從上月清單比對來的（不能就地編輯，編輯了也無處持久化），
  // 「案件內容」展開後多附一行抽查資訊（抽查人員/查證時間/抽查通過時間），這是這個頁籤存在的重點
  function renderLastMonthTable() {
    const tbody = document.getElementById("lastMonthTableBody");
    if (!tbody) return;
    const full = getLastMonthCases();
    // 搜尋只影響「這裡顯示什麼」，不影響 getLastMonthCases() 本身——跟本月案件清單的搜尋是同一套原則，
    // 附近案件/依捷運站分群/路線規劃看到的仍是完整的上月抽查名單，不會因為這裡打了關鍵字就漏掉案件
    const search = state.lastMonthSearch.toLowerCase();
    const list = search
      ? full.filter((c) =>
          (c.address || "").toLowerCase().includes(search) ||
          (c.content || "").toLowerCase().includes(search) ||
          (c.realId || "").toLowerCase().includes(search)
        )
      : full;
    if (!list.length) {
      const msg = !full.length
        ? (state.lastMonthAuditList.length ? "沒有抽查案件" : "尚未匯入抽查清單，請先在上方按「匯入上月案件清單」「匯入抽查清單」")
        : "沒有符合搜尋條件的案件";
      tbody.innerHTML = `<tr><td colspan="8"><div class="empty-state">${msg}</div></td></tr>`;
      return;
    }
    tbody.innerHTML = list.map((c) => {
      const auditBits = [
        c.auditor && `抽查人員：${escapeHtml(c.auditor)}`,
        c.verifiedAt && `查證時間：${escapeHtml(c.verifiedAt)}`,
        c.passedAt && `抽查通過時間：${escapeHtml(c.passedAt)}`,
        c.auditNote && `抽查內容：${escapeHtml(c.auditNote)}`,
      ].filter(Boolean).join("｜");
      const addressCell = c.matched
        ? escapeHtml(c.address)
        : `<span class="hint" style="margin:0">未比對到案件明細，請確認已匯入對應行政區的上月案件清單</span>`;
      return `
        <tr data-id="${escapeHtml(c.id)}" class="${state.rejected.has(c.id) ? "row-rejected" : ""}">
          <td class="case-id-cell" title="${escapeHtml(c.idRaw)}">${escapeHtml(c.realId)}</td>
          <td>${escapeHtml(c.district)}</td>
          <td><span class="badge" title="${escapeHtml(c.category)}" style="background:${colorFor(c.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(c.category))}</span></td>
          <td>${escapeHtml(c.agency)}</td>
          <td>${addressCell}</td>
          <td class="content-cell">${contentCellHtml(c, auditBits ? `<div class="hint" style="margin:4px 0 0">${auditBits}</div>` : "")}</td>
          <td>${escapeHtml(c.createdAt)}</td>
          <td class="heart-td"><div class="heart-cell">
            <span class="heart-btn${state.marked.has(c.id) ? " marked" : ""}" data-case-id="${escapeHtml(c.id)}" title="標記要排路線">${state.marked.has(c.id) ? "❤️" : "🤍"}</span>
            <span class="reject-btn${state.rejected.has(c.id) ? " rejected" : ""}" data-case-id="${escapeHtml(c.id)}" title="${state.rejected.has(c.id) ? "取消否決（恢復正常顯示）" : "標記為暫時否決，不列入抽查考量（只影響顯示，不影響其他功能）"}">🚫</span>
          </div></td>
        </tr>`;
    }).join("");

    tbody.querySelectorAll(".toggle").forEach((t) => {
      t.addEventListener("click", (e) => {
        const cell = e.target.closest(".content-cell");
        cell.classList.toggle("expanded");
        e.target.textContent = cell.classList.contains("expanded") ? "收合內容" : "展開內容";
      });
    });
    tbody.querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", (e) => {
        if (e.target.closest("input, .toggle, .heart-btn, .reject-btn")) return;
        tr.classList.remove("flash-highlight");
        focusCaseOnMap(tr.dataset.id);
      });
    });
    tbody.querySelectorAll(".heart-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleMarked(btn.dataset.caseId);
      });
    });
    tbody.querySelectorAll(".reject-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleRejected(btn.dataset.caseId);
      });
    });
  }

  function renderLastMonthPanel() {
    renderLastMonthTable();
    const el = document.getElementById("lastMonthSummary");
    if (el) {
      el.textContent = (state.lastMonthFullList.length || state.lastMonthAuditList.length)
        ? `上月清單 ${state.lastMonthFullList.length} 筆，抽查清單 ${state.lastMonthAuditList.length} 筆，合併後 ${getLastMonthCases().length} 筆`
        : "";
    }
    broadcastLastMonthView();
  }

  // 就地編輯「發生地址」：改完存回 state.cases 並寫入 localStorage，地圖視窗會透過既有的
  // storage 事件自動同步（跟匯入/清除資料用的是同一套機制），不需要另外通知地圖視窗。
  function addrCellDisplayHtml(c) {
    return `<div class="addr-cell-inner"><span class="addr-edit-btn">✏️</span><span class="addr-text">${escapeHtml(c.address)}</span></div>`;
  }

  function bindAddrCellDisplay(td, c) {
    td.addEventListener("click", (e) => { e.stopPropagation(); startEditAddress(td); });
  }

  function startEditAddress(td) {
    const id = td.dataset.id;
    const c = state.cases.find((x) => x.id === id);
    if (!c) return;
    td.innerHTML = `
      <div style="width:100%">
        <input type="text" class="addr-edit-input" value="${escapeHtml(c.address)}">
        <div style="margin-top:4px;display:flex;gap:6px">
          <button class="ghost addr-save" style="padding:2px 8px;font-size:11px">✓ 儲存</button>
          <button class="ghost addr-cancel" style="padding:2px 8px;font-size:11px">✕ 取消</button>
        </div>
      </div>`;
    const input = td.querySelector(".addr-edit-input");
    input.focus();
    input.select();

    const save = () => {
      const next = input.value.trim();
      if (next && next !== c.address) {
        c.address = next;
        persistCases();
        renderStationGroups(); // 地址變了，分群/定位狀態的判斷依據跟著換
      }
      td.innerHTML = addrCellDisplayHtml(c);
      bindAddrCellDisplay(td, c);
    };
    const cancel = () => {
      td.innerHTML = addrCellDisplayHtml(c);
      bindAddrCellDisplay(td, c);
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") save();
      else if (e.key === "Escape") cancel();
    });
    td.querySelector(".addr-save").addEventListener("click", (e) => { e.stopPropagation(); save(); });
    td.querySelector(".addr-cancel").addEventListener("click", (e) => { e.stopPropagation(); cancel(); });
  }

  function renderPagination(total, totalPages) {
    const indicator = document.getElementById("pageIndicator");
    indicator.textContent = total ? `第 ${state.page} / ${totalPages} 頁（共 ${total} 筆）` : "";
    document.getElementById("btnPrevPage").disabled = state.page <= 1;
    document.getElementById("btnNextPage").disabled = state.page >= totalPages;
  }

  function renderStationGroups() {
    const el = document.getElementById("stationGroupPanel");
    const hintEl = document.getElementById("stationGroupHint");
    // 跟案件清單表格共用同一份 state.filter 的類別排除設定（見 applyCategoryFilters），
    // 只是過濾「這裡顯示什麼」，不影響 state.cases 本身、也不影響地圖/路線規劃（那邊仍是全部已匯入案件）；
    // ♡ 標記只用來排路線，跟這個排除設定是兩件獨立的事。
    // 依捷運站分群本質上是「路線規劃前的輔助工具」（先看某站附近有哪些案件，再決定排哪些進路線），
    // 使用者確認上月抽查案件也要併進來一起分群，這樣規劃某站周邊查訪時能同時看到上月機關抽查過的案件。
    // 但案件項目分析的排除設定（checkbox）跟本月的 AUTO_EXCLUDE_* 是同一類「只該套用在本月資料」的規則
    // ——上月案件是抽查清單已經篩過一輪的固定名單，不該再被本月清單這邊勾選的排除設定連帶影響，
    // 所以排除設定只套在 state.cases，getLastMonthCases() 一律原樣併入、不濾
    const selectedList = applyCategoryFilters(state.cases).concat(getLastMonthCases());

    if (!selectedList.length) {
      el.innerHTML = `<div class="empty-state">${(state.cases.length || getLastMonthCases().length) ? "沒有符合篩選條件的案件" : "尚未匯入任何案件"}</div>`;
      if (hintEl) hintEl.textContent = "";
      return;
    }

    const groups = new Map();
    let missing = 0;
    selectedList.forEach((c) => {
      const loc = state.geocache[cacheKeyFor(c)];
      if (!loc) { missing++; return; }
      const near = nearestStation(loc.lat, loc.lng);
      if (!near) return;
      const key = near.station.name;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ case: c, distance: near.distance, approx: loc.approx, intersection: loc.intersection });
    });

    if (hintEl) hintEl.textContent = missing ? `${selectedList.length} 件，${missing} 件未定位` : `${selectedList.length} 件`;

    if (!groups.size) {
      el.innerHTML = `<div class="empty-state">尚未定位任何案件，請開啟地圖視窗並按「定位案件座標」</div>`;
      return;
    }

    const NEAR_DISTANCE = 500; // 公尺，符合的案件在清單中特別標示
    const sorted = Array.from(groups.entries()).sort((a, b) => b[1].length - a[1].length);
    // 欄位格式跟案件清單表格一樣（案件編號／案件項目／承辦機關／發生地址／案件內容／♡），
    // 只是把「成案時間」換成這裡專屬的「距離」，也不含「行政區」。
    // 真的用 <table> 而不是 CSS grid——每個 .case-row 各自 display:grid 的話，每一列其實是
    // 各自獨立的格線容器，欄寬會各自依那一列自己的內容決定，列與列之間完全不保證對齊
    // （已經在這裡踩過：同一欄不同列時寬時窄，看起來歪七扭八）。<table> + <colgroup> 天生就是
    // 所有列共用同一套欄寬，不會有這個問題，跟案件清單表格用的是同一套邏輯。
    // 每個站群自己一個 <table>（因為要放進各自的 <details> 收合），但全部用同一份 colgroup 比例，
    // 所以看起來還是對齊的；標頭另外用一個沒有 <tbody> 的空表格，套同樣 colgroup 對齊欄位。
    const STATION_COLGROUP = `
      <colgroup>
        <col style="width:13%"><col style="width:13%"><col style="width:9%">
        <col style="width:20%"><col style="width:24%"><col style="width:12%"><col style="width:56px">
      </colgroup>`;
    el.innerHTML =
      (missing ? `<div class="hint" style="margin-bottom:8px">另有 ${missing} 件案件尚未定位，暫不會顯示在下面的分群中。</div>` : "") +
      `<div class="hint" style="margin-bottom:8px"><span class="near-dot"></span>標示為距離捷運站 ${NEAR_DISTANCE} 公尺內的案件；按最後的 ♡ 可標記要排路線、🚫 可標記暫時否決</div>` +
      `<table class="case-table">${STATION_COLGROUP}
        <thead><tr><th>案件編號</th><th>案件項目</th><th>承辦機關</th><th>發生地址</th><th>案件內容</th><th>距離</th><th></th></tr></thead>
      </table>` +
      sorted.map(([station, items]) => {
        const nearCount = items.filter((it) => it.distance <= NEAR_DISTANCE).length;
        return `
        <details class="station-group">
          <summary class="station-group-summary">🚇 ${escapeHtml(station)}　<span style="font-weight:400;color:var(--muted);font-size:12px">(${items.length} 件，${NEAR_DISTANCE}公尺內 ${nearCount} 件)</span></summary>
          <table class="case-table">${STATION_COLGROUP}
            <tbody>
              ${items.sort((a, b) => a.distance - b.distance).map((it) => {
                const c = it.case;
                return `
                <tr class="case-row${it.distance <= NEAR_DISTANCE ? " near-station" : ""}${state.rejected.has(c.id) ? " row-rejected" : ""}" data-case-id="${escapeHtml(c.id)}">
                  <td class="case-id-cell" title="${escapeHtml(c.idRaw)}">${lastMonthTagHtml(c)}${escapeHtml(c.realId || c.id)}</td>
                  <td><span class="badge" title="${escapeHtml(c.category)}" style="background:${colorFor(c.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(c.category))}</span></td>
                  <td>${escapeHtml(c.agency)}</td>
                  <td>${escapeHtml(c.address)}</td>
                  <td class="content-cell">${contentCellHtml(c)}</td>
                  <td class="distance-cell"><span class="dist-value">${Math.round(it.distance)} 公尺</span></td>
                  <td class="heart-td"><div class="heart-cell">
                    <span class="heart-btn${state.marked.has(c.id) ? " marked" : ""}" data-case-id="${escapeHtml(c.id)}" title="標記要排路線">${state.marked.has(c.id) ? "❤️" : "🤍"}</span>
                    <span class="reject-btn${state.rejected.has(c.id) ? " rejected" : ""}" data-case-id="${escapeHtml(c.id)}" title="${state.rejected.has(c.id) ? "取消否決（恢復正常顯示）" : "標記為暫時否決，不列入抽查考量（只影響顯示，不影響其他功能）"}">🚫</span>
                  </div></td>
                </tr>`;
              }).join("")}
            </tbody>
          </table>
        </details>`;
      }).join("");

    el.querySelectorAll(".toggle").forEach((t) => {
      t.addEventListener("click", (e) => {
        e.stopPropagation();
        const cell = e.target.closest(".content-cell");
        cell.classList.toggle("expanded");
        e.target.textContent = cell.classList.contains("expanded") ? "收合內容" : "展開內容";
      });
    });
    el.querySelectorAll(".case-row").forEach((row) => {
      row.addEventListener("click", (e) => {
        if (e.target.closest(".toggle, .heart-btn, .reject-btn")) return;
        focusCaseOnMap(row.dataset.caseId);
      });
    });
    el.querySelectorAll(".heart-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleMarked(btn.dataset.caseId);
      });
    });
    el.querySelectorAll(".reject-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleRejected(btn.dataset.caseId);
      });
    });
  }

  // 地圖視窗點選某案件標記時觸發：以該案件為中心，列出全部已匯入案件中周邊 500 公尺內的其他案件
  // (不限已勾選/標記，方便發現同一區域可以一起排入這趟查訪的其他案件)
  //
  // 這份清單本質上是「輔助/衍生」用途（回答「這附近還有哪些案件」），不是像案件清單、依捷運站
  // 分群那種獨立的工作項目，所以不再自己佔一個左側頁籤（原本這樣做，會跟「回到原本出發的頁籤」
  // 的設計互相打架：從案件清單點過去地圖，點了旁邊案件跳回來，卻要另外切去周邊清單頁籤才看得到
  // 結果，體感很怪）。改成內嵌在案件清單／依捷運站分群／路線規劃這三個會通去地圖的頁籤裡面，
  // 用同一個 class 選取全部現存的容器，不管使用者從哪個頁籤點過去地圖，跳回來就會直接看到。
  const NEARBY_DISTANCE = 500; // 公尺

  function renderNearbyPanel(centerId) {
    const panels = document.querySelectorAll(".nearby-embed-panel");
    if (!panels.length) return;

    // 周邊清單要能跨「本月案件」跟「上月抽查管理」兩個資料集一起找，使用者說看地圖找案件比看文字快，
    // 兩邊案件的座標都存在同一份 state.geocache（用 district+address 當 key，跟資料集無關）
    const combinedCases = state.cases.concat(getLastMonthCases());
    const center = combinedCases.find((c) => c.id === centerId);
    const centerLoc = center && state.geocache[cacheKeyFor(center)];
    let html, hintText;

    if (!center || !centerLoc) {
      html = `<div class="empty-state">此案件尚未定位，無法列出周邊案件</div>`;
      hintText = "";
    } else {
      const nearby = combinedCases
        .filter((c) => c.id !== centerId)
        .map((c) => {
          const loc = state.geocache[cacheKeyFor(c)];
          if (!loc) return null;
          const distance = haversine(centerLoc.lat, centerLoc.lng, loc.lat, loc.lng);
          return distance <= NEARBY_DISTANCE ? { case: c, distance, approx: loc.approx, intersection: loc.intersection } : null;
        })
        .filter(Boolean)
        .sort((a, b) => a.distance - b.distance);

      hintText = `${nearby.length} 件`;
      // 「以案件 #XXX 為中心，N 公尺內案件」這行拿掉了——上面固定的說明文字已經講過這個面板的用途，
      // 使用者反映重複、多餘（見 README）

      // 附近案件現在是固定 300px 寬的窄側欄，放不下原本設計給整頁寬度用的固定像素格線欄位
      // （案件項目/案件編號/地址/距離/♡ 五欄，光欄寬設定就要 500px 以上），改成一列兩行的卡片式排版，
      // 不用對齊欄位，寬度多窄都能正常換行、不會撐爆側欄
      html = !nearby.length
        ? `<div class="empty-state">周邊 ${NEARBY_DISTANCE} 公尺內沒有其他已定位的案件</div>`
        : nearby.map((it) => `
            <div class="mrt-row near-station${state.rejected.has(it.case.id) ? " row-rejected" : ""}" data-case-id="${escapeHtml(it.case.id)}">
              <div class="mrt-row-top">
                <span class="badge" title="${escapeHtml(it.case.category)}" style="background:${colorFor(it.case.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(it.case.category))}</span>
                <span class="mrt-row-distance">${Math.round(it.distance)} 公尺</span>
                <span class="heart-btn${state.marked.has(it.case.id) ? " marked" : ""}" data-case-id="${escapeHtml(it.case.id)}" title="標記要排路線">${state.marked.has(it.case.id) ? "❤️" : "🤍"}</span>
                <span class="reject-btn${state.rejected.has(it.case.id) ? " rejected" : ""}" data-case-id="${escapeHtml(it.case.id)}" title="${state.rejected.has(it.case.id) ? "取消否決（恢復正常顯示）" : "標記為暫時否決，不列入抽查考量（只影響顯示，不影響其他功能）"}">🚫</span>
              </div>
              <div class="mrt-row-addr">${lastMonthTagHtml(it.case)}#${escapeHtml(it.case.realId || it.case.id)}・${escapeHtml(it.case.address)}</div>
            </div>`).join("");
    }

    panels.forEach((el) => {
      el.innerHTML = html;
      el.querySelectorAll(".mrt-row").forEach((row) => {
        row.addEventListener("click", () => focusCaseOnMap(row.dataset.caseId));
      });
      el.querySelectorAll(".heart-btn").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          toggleMarked(btn.dataset.caseId);
        });
      });
      el.querySelectorAll(".reject-btn").forEach((btn) => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          toggleRejected(btn.dataset.caseId);
        });
      });
    });
    document.querySelectorAll(".nearby-embed-hint").forEach((h) => { h.textContent = hintText; });
    if (centerId) document.querySelectorAll(".nearby-embed-details").forEach((d) => { d.open = true; });
  }

  function toggleMarked(id) {
    if (state.marked.has(id)) state.marked.delete(id); else state.marked.add(id);
    if (!state.marked.has(id) && state.routeLockEndId === id) { state.routeLockEndId = null; persistRouteLockEnd(); } // 案件被取消標記，鎖定失去意義，一併解除
    persistMarked();
    updateRouteHint();
    document.querySelectorAll(`.heart-btn[data-case-id="${CSS.escape(id)}"]`).forEach((el) => {
      const marked = state.marked.has(id);
      el.textContent = marked ? "❤️" : "🤍";
      el.classList.toggle("marked", marked);
    });
    renderMarkedList(); // 「路線規劃」頁籤的已標記案件清單要跟著增減列，不能只改圖示——所以整份重畫
  }

  // 「路線規劃」頁籤最上方的「已標記案件」清單：規劃路線前的最後確認名單，可以在這裡直接取消
  // ♡（案件清單/依捷運站分群按 ♡ 標記進來的），也可以點案件跳去地圖核對位置。這裡不需要「附近案件」
  // 側欄（已經不是在看單一案件周邊，而是在看整批要排路線的名單），所以這個頁籤沒有 nearby-embed-panel。
  // 使用者確認上月抽查案件的 ♡ 標記也要能在這裡看到、一起排進路線，跟 state.cases 合併查
  // 使用者反映：標記完案件後可能因故（下雨、選案不合適）拖了幾天才真的出門查，這幾天內
  // 「鄰里無主垃圾清運」「道路散落物或油漬處理」這兩項可能就超過 3 天失效了，但使用者不會
  // 每次都想到要重新按一次「移除3天前垃圾/散落物案件」。這裡只是「提醒」，不會自動幫使用者
  // 移除或取消標記——是否要處理、什麼時候處理，還是使用者自己決定
  function staleMarkedGarbageCases() {
    const cutoff = autoExcludeCutoff();
    return state.cases.filter((c) => state.marked.has(c.id) && isAutoExcludeCase(c, cutoff));
  }

  // 案件項目名稱在這個窄欄位裡截斷用 8 字（比案件清單徽章的 10 字上限更短），單位是「中文字數」，
  // 不是 byte，直接用 JS 字串長度算（中文字在 JS 裡一個字算一個 length，跟 Python 的 len() 一樣）
  function truncateCategoryShort(name) {
    const s = name || "";
    return s.length > 8 ? s.slice(0, 8) + "…" : s;
  }

  // 依案件項目分類列出某個案件清單裡已標記(♡)的筆數／該項目在這份清單裡的總筆數，
  // 用量高的排前面；沒有標記到的項目不列（使用者要求「有選到的案件類別都放上去」，不是固定清單）
  function summarizeMarkedByCategory(cases) {
    const totalByCategory = new Map();
    cases.forEach((c) => totalByCategory.set(c.category, (totalByCategory.get(c.category) || 0) + 1));
    const markedByCategory = new Map();
    cases.forEach((c) => {
      if (!state.marked.has(c.id)) return;
      markedByCategory.set(c.category, (markedByCategory.get(c.category) || 0) + 1);
    });
    return Array.from(markedByCategory.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([category, count]) => ({ category, count, total: totalByCategory.get(category) || count }));
  }

  function summaryRowsHtml(rows) {
    return rows.map(({ category, count, total }) => `
      <div class="side-nav-summary-row" title="${escapeHtml(category)}">
        <span>${escapeHtml(truncateCategoryShort(category))}</span>
        <span class="count">${count}筆/${total}筆</span>
      </div>`).join("");
  }

  // 左側選單下方的「已選」摘要：不管目前切到哪個頁籤都看得到。使用者要求分成「抽查」（本月
  // 案件，state.cases）跟「複查」（上月抽查案件，getLastMonthCases()）兩組各自統計，跟報告
  // 自動更新功能裡「本月4筆抽查＋上月1筆複查」的既有用語一致，不是混在一起算
  function renderSideNavMarkedSummary() {
    const el = document.getElementById("sideNavMarkedSummary");
    if (!el) return;
    const thisMonthRows = summarizeMarkedByCategory(state.cases);
    const lastMonthRows = summarizeMarkedByCategory(getLastMonthCases());
    if (!thisMonthRows.length && !lastMonthRows.length) { el.innerHTML = ""; return; }

    let html = `<div class="side-nav-summary-title">已選</div>`;
    if (thisMonthRows.length) {
      html += `<div class="side-nav-summary-group">抽查</div>` + summaryRowsHtml(thisMonthRows);
    }
    if (lastMonthRows.length) {
      html += `<div class="side-nav-summary-group">複查</div>` + summaryRowsHtml(lastMonthRows);
    }
    el.innerHTML = html;
  }

  function renderMarkedList() {
    renderSideNavMarkedSummary();
    const el = document.getElementById("markedListPanel");
    if (!el) return; // 頁籤還沒切換過去、DOM 還沒建立時，其他地方呼叫 toggleMarked() 不用管這裡
    const hintEl = document.getElementById("markedListHint");
    const markedCases = state.cases.concat(getLastMonthCases()).filter((c) => state.marked.has(c.id));
    if (hintEl) hintEl.textContent = markedCases.length ? `已標記 ${markedCases.length} 筆` : "";

    const staleWarningEl = document.getElementById("staleMarkedWarning");
    if (staleWarningEl) {
      const stale = staleMarkedGarbageCases();
      staleWarningEl.innerHTML = stale.length
        ? `⚠ 已標記案件裡有 ${stale.length} 筆「${AUTO_EXCLUDE_CATEGORIES.join("」「")}」已經超過 ${AUTO_EXCLUDE_DAYS} 天，建議排路線前先移除更新
           <button class="ghost" id="btnStaleWarningRemove" style="margin:0">🗑️ 立即移除</button>`
        : "";
      staleWarningEl.style.display = stale.length ? "flex" : "none";
      const btn = document.getElementById("btnStaleWarningRemove");
      if (btn) btn.addEventListener("click", autoRemoveOldGarbageCases);
    }

    if (!markedCases.length) {
      el.innerHTML = `<div class="empty-state">尚未標記任何案件，請到「本月案件管理」「上月抽查管理」或「依捷運站分群」按 ♡ 標記要排路線的案件</div>`;
      return;
    }

    const MARKED_COLGROUP = `
      <colgroup>
        <col style="width:16%"><col style="width:14%"><col style="width:11%">
        <col style="width:43%"><col style="width:40px"><col style="width:56px">
      </colgroup>`;
    el.innerHTML = `<table class="case-table">${MARKED_COLGROUP}
      <thead><tr><th>案件編號</th><th>案件項目</th><th>承辦機關</th><th>發生地址</th><th title="鎖定為最後一站">📌</th><th></th></tr></thead>
      <tbody>
        ${markedCases.map((c) => `
          <tr data-id="${escapeHtml(c.id)}" class="${state.rejected.has(c.id) ? "row-rejected" : ""}">
            <td class="case-id-cell" title="${escapeHtml(c.idRaw)}">${lastMonthTagHtml(c)}${escapeHtml(c.realId || c.id)}</td>
            <td><span class="badge" title="${escapeHtml(c.category)}" style="background:${colorFor(c.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(c.category))}</span></td>
            <td>${escapeHtml(c.agency)}</td>
            <td>${escapeHtml(c.address)}</td>
            <td class="lock-td"><span class="lock-btn${state.routeLockEndId === c.id ? " locked" : ""}" data-case-id="${escapeHtml(c.id)}" title="鎖定為最後一站（規劃路線時固定排最後，其餘案件仍會自動重新排序）">📌</span></td>
            <td class="heart-td"><div class="heart-cell"><span class="heart-btn marked" data-case-id="${escapeHtml(c.id)}" title="取消標記">❤️</span></div></td>
          </tr>`).join("")}
      </tbody>
    </table>`;

    el.querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", (e) => {
        if (e.target.closest(".heart-btn, .lock-btn")) return;
        tr.classList.remove("flash-highlight");
        focusCaseOnMap(tr.dataset.id);
      });
    });
    el.querySelectorAll(".lock-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleRouteLockEnd(btn.dataset.caseId);
      });
    });
    el.querySelectorAll(".heart-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleMarked(btn.dataset.caseId);
      });
    });
  }

  // 「暫時否決、不列入抽查考量」：純粹是顯示層的樣式標記（背景變白、文字變淺灰），
  // 只在畫面上就地切換 class，不碰 state.cases、不影響篩選/地圖/路線規劃等其他機制，
  // 避免跟既有的排除設定、♡ 標記路線互相干擾。
  function toggleRejected(id) {
    if (state.rejected.has(id)) state.rejected.delete(id); else state.rejected.add(id);
    persistRejected();
    const rejected = state.rejected.has(id);
    document.querySelectorAll(`[data-id="${CSS.escape(id)}"], [data-case-id="${CSS.escape(id)}"]`).forEach((el) => {
      el.classList.toggle("row-rejected", el.matches("tr, .mrt-row") ? rejected : el.classList.contains("row-rejected"));
    });
    document.querySelectorAll(`.reject-btn[data-case-id="${CSS.escape(id)}"]`).forEach((el) => {
      el.classList.toggle("rejected", rejected);
      el.title = rejected ? "取消否決（恢復正常顯示）" : "標記為暫時否決，不列入抽查考量（只影響顯示，不影響其他功能）";
    });
  }

  // 判斷一筆案件是否符合「查驗共識」：這兩項案件項目、且成案時間為 AUTO_EXCLUDE_DAYS 天前。
  // 匯入時的排除勾選、案件清單的「移除」按鈕共用同一套判斷，避免兩處邏輯寫兩次、之後改天數只改一處
  function isAutoExcludeCase(c, cutoff) {
    if (!AUTO_EXCLUDE_CATEGORIES.includes(c.category)) return false;
    const createdAt = parseRocDateTime(c.createdAt);
    return createdAt && createdAt < cutoff;
  }
  function autoExcludeCutoff() {
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - AUTO_EXCLUDE_DAYS);
    return cutoff;
  }

  // 查驗共識：「鄰里無主垃圾清運」「道路散落物或油漬處理」只查 3 天內成案的，3 天前的直接不列入
  // 考量。使用者確認要「真的從資料裡移除」（不是否決標記）——案件量大時否決標記還留在清單裡、只是
  // 變淡，一樣要捲過去，希望清單本身變短。原本有一個「復原」按鈕可以把剛移除的一批找回來，使用者
  // 確認這兩項一旦超過 3 天就真的沒有查驗意義了、不會想復原，已經拿掉這個功能，移除就是真的移除。
  // 匯入那一刻就想先排除掉的話，見「資料匯入」頁籤的排除勾選框（見 handleFiles）
  function autoRemoveOldGarbageCases() {
    const cutoff = autoExcludeCutoff();
    const targets = state.cases.filter((c) => isAutoExcludeCase(c, cutoff));
    if (!targets.length) {
      alert(`沒有符合條件的案件（「${AUTO_EXCLUDE_CATEGORIES.join("」「")}」且成案時間為 ${AUTO_EXCLUDE_DAYS} 天前的案件）。`);
      return;
    }
    if (!confirm(`將從資料中移除 ${targets.length} 筆「${AUTO_EXCLUDE_CATEGORIES.join("」「")}」（${AUTO_EXCLUDE_DAYS} 天前成案）案件，確定嗎？`)) return;
    const targetIds = new Set(targets.map((c) => c.id));
    state.cases = state.cases.filter((c) => !targetIds.has(c.id));
    targetIds.forEach((id) => { state.marked.delete(id); state.rejected.delete(id); });
    persistCases(); persistMarked(); persistRejected();
    renderAll(); // 一次改動很多筆，直接整份重畫比逐筆補 DOM 簡單可靠
    alert(`已移除 ${targets.length} 筆案件。`);
  }

  function updateRouteHint() {
    const hint = document.getElementById("routeHint");
    if (hint) hint.textContent = state.marked.size ? `已標記 ${state.marked.size} 筆` : "";
  }

  // shouldOpenMap: 使用者主動按「規劃路線」時才彈開/聚焦地圖視窗；重新整理頁面後靜靜還原顯示時不用。
  // 使用者確認上月抽查案件的 ♡ 標記也要真的排進路線，不是只在上面的清單看得到，跟 state.cases 合併查
  function refreshRoutePanel(shouldOpenMap) {
    const markedCases = state.cases.concat(getLastMonthCases()).filter((c) => state.marked.has(c.id));
    if (!markedCases.length) {
      document.getElementById("routePanel").innerHTML = "";
      return;
    }

    const points = [];
    let unlocated = 0;
    markedCases.forEach((c) => {
      const loc = state.geocache[cacheKeyFor(c)];
      if (loc) points.push({ id: c.id, lat: loc.lat, lng: loc.lng, case: c });
      else unlocated++;
    });

    if (points.length < 2) {
      document.getElementById("routePanel").innerHTML =
        `<div class="empty-state">已標記且已定位的案件只有 ${points.length} 筆，至少需要 2 筆才能規劃路線（未定位的案件請先到地圖視窗按「定位案件座標」）。</div>`;
      return;
    }

    const startName = document.getElementById("routeStartStation").value;
    const startStation = startName ? (window.MRT_STATIONS || []).find((s) => s.name === startName) : null;

    // 使用者鎖定的「最後一站」不一定在目前這批已標記且已定位的案件裡（例如剛好取消標記或還沒定位），
    // 找不到就當沒鎖定，靜靜退回原本全自動排序，不用跳錯誤訊息
    let orderedPoints;
    if (startStation) {
      const withStart = [{ id: null, lat: startStation.lat, lng: startStation.lng, isStation: true }, ...points];
      const endIdx = state.routeLockEndId ? withStart.findIndex((p) => p.id === state.routeLockEndId) : -1;
      orderedPoints = planRouteOrder(withStart, 0, endIdx >= 0 ? endIdx : null).map((i) => withStart[i]).slice(1);
    } else {
      const endIdx = state.routeLockEndId ? points.findIndex((p) => p.id === state.routeLockEndId) : -1;
      orderedPoints = planRouteOrder(points, null, endIdx >= 0 ? endIdx : null).map((i) => points[i]);
    }

    renderRoutePanel(orderedPoints, unlocated, startStation);
    lastRoute = { startStation, order: orderedPoints, unlocated };
    persistRouteOrder(orderedPoints, startStation);
    persistRouteStart(startName);
    if (shouldOpenMap) openMapWindow();
  }

  // 把目前的路線順序寫進 localStorage，讓地圖視窗（監聽 LS_ROUTE 的 storage 事件）同步重畫。
  // 拖曳手動調整順序、跟按「規劃路線」自動排序，最後都是呼叫這個函式同步地圖，寫的格式一樣
  function persistRouteOrder(order, startStation) {
    localStorage.setItem(LS_ROUTE, JSON.stringify({
      order: order.map((p) => p.id),
      start: startStation ? { name: startStation.name, lat: startStation.lat, lng: startStation.lng } : null,
      ts: Date.now(),
    }));
  }

  // 使用者拖曳排序後呼叫：只調整順序、重新算距離跟編號，不重新跑自動最佳化演算法——
  // 拖到哪個位置就照那個位置顯示，直到使用者再按一次「規劃路線」才會整批重新自動排序
  function reorderRoute(fromIndex, toIndex) {
    if (!lastRoute || fromIndex === toIndex) return;
    const order = lastRoute.order.slice();
    const [moved] = order.splice(fromIndex, 1);
    order.splice(toIndex, 0, moved);
    lastRoute.order = order;
    renderRoutePanel(order, lastRoute.unlocated || 0, lastRoute.startStation);
    persistRouteOrder(order, lastRoute.startStation);
  }

  function planAndShowRoute() {
    if (!state.marked.size) { alert("請先到「本月案件管理」「上月抽查管理」或「依捷運站分群」按 ♡ 標記要排路線的案件。"); return; }
    refreshRoutePanel(true);
  }

  function renderRoutePanel(order, unlocatedCount, startStation) {
    const el = document.getElementById("routePanel");
    let prev = startStation ? { lat: startStation.lat, lng: startStation.lng } : null;

    const startRow = startStation ? `
      <div class="route-row" style="cursor:default">
        <div class="route-step" style="background:var(--good)">🚇</div>
        <div></div>
        <div></div>
        <div style="font-weight:600">起點：${escapeHtml(startStation.name)}</div>
        <div></div>
        <div></div>
      </div>` : "";

    const rows = order.map((p, i) => {
      let legHtml = "";
      if (prev) {
        const d = haversine(prev.lat, prev.lng, p.lat, p.lng);
        legHtml = `<div class="route-leg">↓ ${Math.round(d)} 公尺</div>`;
      }
      prev = p;
      return `
        <div class="route-row" draggable="true" data-order-index="${i}" data-case-id="${escapeHtml(p.id)}" title="拖曳左側編號可調整順序">
          <div class="route-step">${i + 1}</div>
          <div><span class="badge" title="${escapeHtml(p.case.category)}" style="background:${colorFor(p.case.category)}22;color:var(--ink)">${escapeHtml(truncateCategory(p.case.category))}</span></div>
          <div class="case-id">${lastMonthTagHtml(p.case)}${state.routeLockEndId === p.id ? '<span title="鎖定為最後一站">📌</span> ' : ""}#${escapeHtml(p.case.realId || p.id)}</div>
          <div class="addr-cell" title="${escapeHtml(p.case.address)}">${escapeHtml(p.case.address)}</div>
          <div class="content-cell">${contentCellHtml(p.case)}</div>
          <span class="heart-btn marked" data-case-id="${escapeHtml(p.id)}" title="取消標記（會立刻重新規劃路線，不含這一站）">❤️</span>
        </div>
        ${legHtml}`;
    }).join("");

    el.innerHTML =
      startRow + rows +
      (unlocatedCount ? `<div class="hint" style="margin-top:8px;color:#c9821a">另有 ${unlocatedCount} 筆已標記但尚未定位，未列入路線。</div>` : "");

    el.querySelectorAll(".route-row[data-case-id]").forEach((row) => {
      row.addEventListener("click", (e) => {
        if (e.target.closest(".toggle, .heart-btn")) return; // 展開內容/取消標記不該連動跳去地圖
        focusCaseOnMap(row.dataset.caseId);
      });
    });
    el.querySelectorAll(".toggle").forEach((t) => {
      t.addEventListener("click", (e) => {
        e.stopPropagation();
        const cell = e.target.closest(".content-cell");
        cell.classList.toggle("expanded");
        e.target.textContent = cell.classList.contains("expanded") ? "收合內容" : "展開內容";
      });
    });
    // 使用者反映規劃完看到實際路線後，可能臨時反悔某一站不想去了，希望能直接在這裡取消，
    // 不用再切回上面的「已標記案件」清單找。跟其他地方的 ♡ 不同：這裡取消標記後立刻重新規劃路線
    // （少了那一站），因為使用者已經在看「查訪路線」這個結果、就是想看少一站之後的樣子，
    // 跟案件清單/依捷運站分群等別的地方單純標記、不自動連動規劃是刻意的不同（那些地方標記時
    // 使用者通常還在瀏覽/篩選案件，不是在確認最終路線，不該無預警幫他重新規劃、跳地圖畫面）
    el.querySelectorAll(".route-row .heart-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        toggleMarked(btn.dataset.caseId);
        refreshRoutePanel(false);
      });
    });

    // 拖曳手動調整順序（原生 HTML5 drag-and-drop，不用外部套件）。從 ❤️/展開內容 上面按住不算
    // 拖曳起點，避免使用者只是想點這兩個按鈕卻不小心觸發拖曳
    let dragFromIndex = null;
    el.querySelectorAll(".route-row[data-order-index]").forEach((row) => {
      row.addEventListener("dragstart", (e) => {
        if (e.target.closest(".toggle, .heart-btn")) { e.preventDefault(); return; }
        dragFromIndex = Number(row.dataset.orderIndex);
        row.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", String(dragFromIndex));
      });
      row.addEventListener("dragend", () => {
        row.classList.remove("dragging");
        el.querySelectorAll(".route-row.drag-over-top, .route-row.drag-over-bottom").forEach((r) => r.classList.remove("drag-over-top", "drag-over-bottom"));
      });
      row.addEventListener("dragover", (e) => {
        if (dragFromIndex == null) return;
        e.preventDefault();
        const rect = row.getBoundingClientRect();
        const before = e.clientY - rect.top < rect.height / 2;
        row.classList.toggle("drag-over-top", before);
        row.classList.toggle("drag-over-bottom", !before);
      });
      row.addEventListener("dragleave", () => {
        row.classList.remove("drag-over-top", "drag-over-bottom");
      });
      row.addEventListener("drop", (e) => {
        e.preventDefault();
        const before = row.classList.contains("drag-over-top");
        row.classList.remove("drag-over-top", "drag-over-bottom");
        if (dragFromIndex == null) return;
        let toIndex = Number(row.dataset.orderIndex);
        if (!before) toIndex += 1;
        if (dragFromIndex < toIndex) toIndex -= 1; // 從陣列移除來源那筆後，後面的目標位置索引要往前補一格
        reorderRoute(dragFromIndex, toIndex);
        dragFromIndex = null;
      });
    });
  }

  function renderAll() {
    renderStatCards();
    renderBreakdown("breakdownCategory", state.cases, "category", "category");
    renderBreakdown("breakdownAgency", state.cases, "agency", "agency");
    renderTable();
    renderStationGroups();
    updateRouteHint();
    renderMarkedList();
  }
  const render = renderAll;

  // ---------- CSV 匯出 (供 Google 我的地圖匯入) ----------
  // 匯出目前清單 (套用篩選/搜尋條件，不限分頁) 為 CSV
  function exportCsv() {
    const list = applyFilters();
    if (!list.length) { alert("目前沒有任何案件可以匯出。"); return; }
    const header = ["案件編號", "案件項目", "承辦機關", "行政區", "發生地址", "案件內容", "案件處理", "狀態", "成案時間"];
    const rows = list.map((c) => [c.id, c.category, c.agency, c.district, "台北市" + c.address, c.content, c.handling, c.status, c.createdAt]);
    const csv = [header, ...rows].map((r) => r.map((v) => `"${String(v || "").replace(/"/g, '""')}"`).join(",")).join("\r\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const ts = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `案件清單_${ts}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function exportRoute() {
    if (!lastRoute || !lastRoute.order.length) { alert("請先按「規劃路線」產生路線後再匯出。"); return; }
    const { startStation, order } = lastRoute;
    const header = ["順序", "案件編號", "案件項目", "行政區", "發生地址", "案件內容", "案件處理", "到下一站距離(公尺)"];
    const rows = [];
    if (startStation) {
      const d = order.length ? Math.round(haversine(startStation.lat, startStation.lng, order[0].lat, order[0].lng)) : "";
      rows.push(["起點", "", "", "", `捷運${startStation.name}`, "", "", d]);
    }
    order.forEach((p, i) => {
      const toNext = i < order.length - 1 ? Math.round(haversine(p.lat, p.lng, order[i + 1].lat, order[i + 1].lng)) : "";
      const idLabel = (p.case.realId || p.id) + (p.case.id.indexOf("抽") === 0 ? "（上月抽查）" : "");
      rows.push([i + 1, idLabel, p.case.category, p.case.district, "台北市" + p.case.address, p.case.content, p.case.handling, toNext]);
    });
    const csv = [header, ...rows].map((r) => r.map((v) => `"${String(v == null ? "" : v).replace(/"/g, '""')}"`).join(",")).join("\r\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const ts = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `查訪路線_${ts}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // 讀取地圖截圖（不管是從檔案選的還是從剪貼簿貼的），轉成 Word 匯出要用的格式：
  // 二進位資料 (Uint8Array) 給 docx 的 ImageRun，加上原始寬高比例，匯出時才能等比例縮放
  async function loadRouteMapImageFile(file, label) {
    try {
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      const type = file.type.includes("png") ? "png" : "jpg";
      const dims = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => { resolve({ width: img.naturalWidth, height: img.naturalHeight }); URL.revokeObjectURL(img.src); };
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
      });
      routeMapImageData = { bytes, type, width: dims.width, height: dims.height };
      document.getElementById("routeMapImageStatus").textContent = `已選擇：${label || file.name || "剪貼簿圖片"}`;
    } catch (err) {
      alert("這張圖片讀取失敗，請確認是圖片檔（png/jpg）。");
    }
  }

  function handleRouteMapImageChange(e) {
    const file = e.target.files[0];
    e.target.value = "";
    if (file) loadRouteMapImageFile(file);
  }

  // Windows「剪取工具」(Win+Shift+S) 截圖後預設就是複製到剪貼簿，不用先存檔——所以除了選檔案，
  // 也支援直接按 Ctrl+V 貼上剪貼簿裡的圖片。只在剪貼簿內容「是圖片」時才處理、並擋掉預設貼上行為，
  // 貼的是文字（例如在搜尋框、地址編輯欄位貼字）時完全不受影響，照瀏覽器原本的行為就好
  document.addEventListener("paste", (e) => {
    const items = (e.clipboardData && e.clipboardData.items) || [];
    const imageItem = Array.from(items).find((it) => it.type && it.type.startsWith("image/"));
    if (!imageItem) return;
    const file = imageItem.getAsFile();
    if (!file) return;
    e.preventDefault();
    loadRouteMapImageFile(file, "剪貼簿截圖");
  });

  // 匯出查訪路線的 Word 文件：每筆案件一個區塊（順序/案件編號/承辦機關/距離、項目、地址、內容），
  // 依規劃好的順序排列，方便帶出去現場查訪時列印或用平板對照。用 docx 這個 JS 套件在瀏覽器端
  // 直接產生 .docx，不需要伺服器；最上方可選擇貼一張使用者自己截圖的地圖畫面
  async function exportRouteWord() {
    if (!lastRoute || !lastRoute.order.length) { alert("請先按「規劃路線」產生路線後再匯出。"); return; }
    if (!window.docx) { alert("Word 匯出功能所需的元件還在載入，請稍後再試一次。"); return; }
    const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType } = window.docx;
    const { startStation, order } = lastRoute;
    // 使用者要求檔名要有產出的日期+時間（同一天規劃、匯出好幾次路線時才不會互相覆蓋）；
    // 檔名不能有冒號，跟文件內文顯示用的日期格式分開處理——文件內文人看的格式維持易讀，
    // 檔名另外用底線接時間，本地時間（不是 toISOString() 的 UTC，差 8 小時對使用者來說會誤導）
    const now = new Date();
    const pad2 = (n) => String(n).padStart(2, "0");
    const displayTs = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())} ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
    const ts = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}`;

    const children = [];
    if (routeMapImageData) {
      const maxWidth = 600, maxHeight = 480;
      let w = routeMapImageData.width, h = routeMapImageData.height;
      const scale = Math.min(1, maxWidth / w, maxHeight / h);
      w = Math.round(w * scale);
      h = Math.round(h * scale);
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new ImageRun({ type: routeMapImageData.type, data: routeMapImageData.bytes, transformation: { width: w, height: h } })],
      }));
    }
    children.push(
      new Paragraph({ text: "派工案件抽查路線", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `匯出日期：${displayTs}　共 ${order.length} 件${startStation ? `　起點：捷運${startStation.name}` : ""}`, size: 20, color: "666666" })],
        spacing: { after: 300 },
      })
    );

    order.forEach((p, i) => {
      const toNext = i < order.length - 1 ? Math.round(haversine(p.lat, p.lng, order[i + 1].lat, order[i + 1].lng)) : null;
      const c = p.case;
      const headerRuns = [
        new TextRun({ text: `${i + 1}. `, bold: true, size: 26 }),
        new TextRun({ text: `案件編號：${c.realId || c.id}${c.id.indexOf("抽") === 0 ? "（上月抽查）" : ""}`, bold: true, size: 24 }),
        new TextRun({ text: `　承辦機關：${c.agency || "（未填）"}`, size: 24 }),
      ];
      if (toNext != null) headerRuns.push(new TextRun({ text: `　距下一站：約 ${toNext} 公尺`, size: 24, color: "1d6fd6" }));
      children.push(
        new Paragraph({
          spacing: { before: i === 0 ? 0 : 200, after: 60 },
          border: { top: { style: "single", size: 6, color: "CCCCCC" } },
          children: headerRuns,
        }),
        new Paragraph({ children: [new TextRun({ text: `項目：${c.category}` })] }),
        new Paragraph({ children: [new TextRun({ text: `地址：台北市${c.address}` })] }),
        new Paragraph({ children: [new TextRun({ text: `內容：${c.content}` })] }),
        ...(c.handling ? [new Paragraph({ children: [new TextRun({ text: `處理：${c.handling}` })] })] : [])
      );
    });

    const doc = new Document({ sections: [{ properties: {}, children }] });
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `查訪路線_${ts}.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // ---------- 地理編碼設定 (Google API 金鑰選填，讓地圖視窗改用較快的 Google Geocoding API) ----------
  function loadGeocodeSettings() {
    try { return JSON.parse(localStorage.getItem(LS_SETTINGS) || "{}"); } catch (e) { return {}; }
  }

  function activeGeocodeProvider(settings) {
    const s = settings || loadGeocodeSettings();
    if (s.geocodeProvider === "osm") return "osm";
    return s.googleApiKey ? "google" : "osm"; // "auto" 跟 "google" 沒填金鑰時都退回 osm
  }

  function renderGeocodeSettingsHint() {
    const s = loadGeocodeSettings();
    const provider = activeGeocodeProvider(s);
    const geoLabel = provider === "google" ? "Google（較快）" : "OpenStreetMap（免費，較慢）";
    const tileLabel = s.tileProvider === "google" && s.googleApiKey ? "Google" : "OpenStreetMap";
    document.getElementById("geocodeSettingsHint").textContent = `目前地理編碼：${geoLabel}／地圖底圖：${tileLabel}`;
  }

  function initGeocodeSettingsUi() {
    const s = loadGeocodeSettings();
    document.getElementById("geocodeProvider").value = s.geocodeProvider || "auto";
    document.getElementById("tileProvider").value = s.tileProvider || "osm";
    document.getElementById("googleApiKeyInput").value = s.googleApiKey || "";
    renderGeocodeSettingsHint();
    document.getElementById("btnSaveGeocodeSettings").addEventListener("click", () => {
      const next = {
        geocodeProvider: document.getElementById("geocodeProvider").value,
        tileProvider: document.getElementById("tileProvider").value,
        googleApiKey: document.getElementById("googleApiKeyInput").value.trim(),
      };
      localStorage.setItem(LS_SETTINGS, JSON.stringify(next));
      renderGeocodeSettingsHint();
      alert("已儲存。開啟或已開啟的地圖視窗會套用新設定（底圖切換若失敗會自動退回 OpenStreetMap）。");
    });
  }

  // ---------- 事件綁定 ----------
  function bindEvents() {
    document.querySelectorAll(".side-nav-item").forEach((btn) => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });

    document.getElementById("btnUpload").addEventListener("click", () => document.getElementById("fileInput").click());
    document.getElementById("fileInput").addEventListener("change", (e) => {
      const mode = document.getElementById("importMode").value;
      handleFiles(e.target.files, mode);
      e.target.value = "";
    });
    document.getElementById("btnClearData").addEventListener("click", () => {
      if (!state.cases.length) return;
      if (!confirm("確定要清除目前所有已匯入的案件資料嗎？此動作無法復原。")) return;
      state.cases = [];
      state.geocache = {};
      state.marked = new Set();
      state.rejected = new Set();
      state.filter = emptyFilter();
      state.page = 1;
      persistCases(); persistGeocache(); persistMarked(); persistRejected();
      lastRoute = null;
      localStorage.removeItem(LS_ROUTE);
      document.getElementById("searchBox").value = "";
      document.getElementById("routePanel").innerHTML = "";
      renderAll();
    });

    document.getElementById("btnUploadLastMonthFull").addEventListener("click", () => document.getElementById("lastMonthFullInput").click());
    document.getElementById("lastMonthFullInput").addEventListener("change", (e) => {
      handleLastMonthFullFiles(e.target.files);
      e.target.value = "";
    });
    document.getElementById("btnUploadLastMonthAudit").addEventListener("click", () => document.getElementById("lastMonthAuditInput").click());
    document.getElementById("lastMonthAuditInput").addEventListener("change", (e) => {
      handleLastMonthAuditFiles(e.target.files);
      e.target.value = "";
    });
    document.getElementById("btnClearLastMonth").addEventListener("click", () => {
      if (!state.lastMonthFullList.length && !state.lastMonthAuditList.length) return;
      if (!confirm("確定要清除上月抽查管理的資料嗎？此動作無法復原。")) return;
      state.lastMonthFullList = [];
      state.lastMonthAuditList = [];
      persistLastMonthFullList();
      persistLastMonthAuditList();
      renderLastMonthPanel();
    });
    document.getElementById("lastMonthSearchBox").addEventListener("input", (e) => {
      state.lastMonthSearch = e.target.value.trim();
      renderLastMonthTable();
    });
    document.getElementById("btnClearLastMonthFilter").addEventListener("click", () => {
      state.lastMonthSearch = "";
      document.getElementById("lastMonthSearchBox").value = "";
      renderLastMonthTable();
    });

    document.getElementById("searchBox").addEventListener("input", (e) => {
      state.filter.search = e.target.value.trim();
      state.page = 1;
      renderTable();
    });
    document.getElementById("btnClearFilter").addEventListener("click", () => {
      state.filter = emptyFilter();
      document.getElementById("searchBox").value = "";
      state.page = 1;
      renderAll();
    });
    document.getElementById("btnAutoRemoveOld").addEventListener("click", autoRemoveOldGarbageCases);
    document.getElementById("excludeOldOnImport").checked = localStorage.getItem(LS_EXCLUDE_ON_IMPORT) === "1";
    document.getElementById("excludeOldOnImport").addEventListener("change", (e) => {
      localStorage.setItem(LS_EXCLUDE_ON_IMPORT, e.target.checked ? "1" : "0");
    });
    document.getElementById("btnPrevPage").addEventListener("click", () => {
      state.page--;
      renderTable();
    });
    document.getElementById("btnNextPage").addEventListener("click", () => {
      state.page++;
      renderTable();
    });
    document.getElementById("pageSizeSelect").addEventListener("change", (e) => {
      state.pageSize = parseInt(e.target.value, 10);
      state.page = 1;
      localStorage.setItem(LS_PAGE_SIZE, String(state.pageSize));
      renderTable();
    });
    document.getElementById("btnExportCsv").addEventListener("click", exportCsv);
    document.getElementById("btnOpenMap").addEventListener("click", openMapWindow);
    document.getElementById("btnPlanRoute").addEventListener("click", planAndShowRoute);
    document.getElementById("btnExportRoute").addEventListener("click", exportRoute);
    document.getElementById("btnExportRouteWord").addEventListener("click", exportRouteWord);
    document.getElementById("btnChooseRouteMapImage").addEventListener("click", () => {
      document.getElementById("routeMapImageInput").click();
    });
    document.getElementById("routeMapImageInput").addEventListener("change", handleRouteMapImageChange);
    document.getElementById("btnClearMarked").addEventListener("click", () => {
      if (!state.marked.size) return;
      if (!confirm("確定要清空所有已標記的路線案件嗎？")) return;
      state.marked = new Set();
      persistMarked();
      lastRoute = null;
      localStorage.removeItem(LS_ROUTE);
      document.getElementById("routePanel").innerHTML = "";
      renderStationGroups();
      updateRouteHint();
      renderMarkedList();
    });
    document.querySelectorAll("th[data-sort]").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.dataset.sort;
        if (state.sort.key === key) state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
        else { state.sort.key = key; state.sort.dir = "asc"; }
        renderTable();
      });
    });
  }

  // 地圖視窗完成定位後會寫入 LS_GEOCACHE，這裡即時同步更新分群結果。
  // 案件多的時候地圖視窗會自動連續定位上百筆，等於每 1 秒左右就寫一次 LS_GEOCACHE，
  // 如果每次都馬上重畫分群清單（要對每筆案件重新算最近捷運站、重建整塊 HTML），
  // 主視窗跟地圖視窗共用同一個瀏覽器分頁行程，會卡到使用者當下在主視窗的操作（例如點編輯地址的鉛筆）。
  // 改成節流：資料本身（state.geocache）馬上更新，但畫面最多每 1.5 秒才重畫一次
  let stationGroupsRenderTimer = null;
  window.addEventListener("storage", (e) => {
    if (e.key === LS_GEOCACHE) {
      try { state.geocache = JSON.parse(localStorage.getItem(LS_GEOCACHE) || "{}"); } catch (err) { state.geocache = {}; }
      if (!stationGroupsRenderTimer) {
        stationGroupsRenderTimer = setTimeout(() => {
          stationGroupsRenderTimer = null;
          renderStationGroups();
        }, 1500);
      }
    } else if (e.key === LS_FOCUS) {
      // 地圖視窗有人點了某個標記，跳回清單對應那一列，並列出該案件周邊 500 公尺內的其他案件
      try {
        const req = JSON.parse(e.newValue || "null");
        if (req && req.id) { revealCaseByOrigin(req.id); renderNearbyPanel(req.id); }
      } catch (err) {}
    } else if (e.key === LS_MARKED) {
      // 使用者現在也可以直接在地圖的資訊卡按 ❤️，主視窗這邊要跟著重畫（案件清單/分群/已標記清單的圖示跟筆數）
      try { state.marked = new Set(JSON.parse(localStorage.getItem(LS_MARKED) || "[]")); } catch (err) { state.marked = new Set(); }
      renderAll();
      renderLastMonthPanel();
    } else if (e.key === LS_REJECTED) {
      // 同樣道理：地圖資訊卡按 🚫 否決，主視窗清單也要跟著淡化對應那幾列
      try { state.rejected = new Set(JSON.parse(localStorage.getItem(LS_REJECTED) || "[]")); } catch (err) { state.rejected = new Set(); }
      renderAll();
      renderLastMonthPanel();
    } else if (e.key === LS_CATEGORY_FILTER) {
      // 使用者也可以直接點地圖下方圖例的圓點篩選（跟主視窗點色點同一套邏輯），主視窗這邊要跟著把
      // 「依案件項目」清單的勾選狀態同步過來，不然會變成兩邊看到的排除設定不一致
      try { state.filter.category = new Set(JSON.parse(localStorage.getItem(LS_CATEGORY_FILTER) || "[]")); } catch (err) { state.filter.category = new Set(); }
      renderAll();
    }
  });

  // ---------- 初始化 ----------
  function populateRouteStartOptions() {
    const sel = document.getElementById("routeStartStation");
    const stations = (window.MRT_STATIONS || []).slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant"));
    stations.forEach((st) => {
      const opt = document.createElement("option");
      opt.value = st.name;
      opt.textContent = st.name;
      sel.appendChild(opt);
    });
    const saved = localStorage.getItem(LS_ROUTE_START);
    if (saved && stations.some((st) => st.name === saved)) sel.value = saved;
  }

  document.addEventListener("DOMContentLoaded", () => {
    loadState();
    bindEvents();
    switchTab(state.activeTab);
    initGeocodeSettingsUi();
    populateRouteStartOptions();
    document.getElementById("pageSizeSelect").value = String(state.pageSize);
    renderAll();
    renderLastMonthPanel();
    if (state.marked.size >= 2) refreshRoutePanel(false); // 重新整理後靜靜還原之前規劃好的路線顯示，不彈開地圖視窗
  });
})();
