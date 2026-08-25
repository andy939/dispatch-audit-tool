/* 派工案件地圖視窗 - 讀取主視窗透過 localStorage 分享的案件/地理編碼快取資料 */
(function () {
  "use strict";

  const LS_CASES = "dispatch_cases_v1";
  const LS_GEOCACHE = "dispatch_geocache_v1";
  const LS_FOCUS = "dispatch_focus_v1"; // 雙向「跳到地圖/跳回清單」訊號，與主視窗共用同一個 key
  const LS_ROUTE = "dispatch_route_v1"; // 主視窗規劃好的路線 (依訪查順序排列的案件編號)
  const LS_SETTINGS = "dispatch_settings_v1"; // 地理編碼設定 (Google API 金鑰、服務選擇)，主視窗設定完會透過這個 key 同步過來
  const LS_REJECTED = "dispatch_rejected_v1"; // 主視窗標記「暫時否決」的案件，這裡同步過來把點位淡化顯示
  const LS_MARKED = "dispatch_marked_v1"; // 標記要排路線的案件；資訊卡的 ❤️ 也能直接改，跟主視窗雙向同步
  const LS_LASTMONTH_CASES = "dispatch_lastmonth_view_v1"; // 「上月抽查管理」合併後的清單，主視窗單向廣播過來，這裡只讀不寫

  // 使用者要求參考 Google 地圖的配色風格；Google 地圖本身的圖標配色沒有公開對外的固定色碼表，
  // 改採 Google 自家 Material Design 色盤（Google 產品一致採用的標準色系，同樣是這種飽和、乾淨
  // 的視覺風格），取 12 色的 500 色階，避開灰/黑（跟「捷運站」圖例的灰色點混在一起分不清楚）。
  // 跟 app.js 要用同一份，兩邊必須完全一樣。
  const COLORS = ["#F44336", "#E91E63", "#9C27B0", "#673AB7", "#3F51B5", "#2196F3",
    "#009688", "#4CAF50", "#FFC107", "#FF9800", "#FF5722", "#795548"];

  // 這兩項是抽查流程裡最後才查驗的項目，使用者要求固定用灰階區分開來，不要跟著雜湊自動配色。
  // 跟 app.js 要用同一份，兩邊必須完全一樣。
  const FIXED_COLORS = { "鄰里無主垃圾清運": "#cccccc", "道路散落物或油漬處理": "#4a4a4a" };

  let cases = [];
  let geocache = {};
  let rejected = new Set();
  let marked = new Set();
  let lastMonthCases = [];

  let map, mrtLayer, caseLayer, routeLayer;
  const markersById = new Map();

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

  // 全形數字/英文/符號轉半形，避免地理編碼查無此地址（例如「明德路２１０號」→「明德路210號」）
  function toHalfWidth(s) {
    return (s || "")
      .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
      .replace(/　/g, " ");
  }

  // 去掉門牌號（含樓層/範圍/備註），保留到路/街/巷/弄，供精確地址查無結果時做「約略到路」的備援查詢
  // 優先切在最後一個「N巷」或「N弄」之後（例如「大業路527巷57-133(中央北路四段)」→「大業路527巷」），
  // 因為很多案件地址是「巷+號碼區間」這種沒有「號」字的格式，光靠找「號」切不出來。
  function roadOnly(address) {
    let cut = -1, cutEnd = -1;
    const re = /\d+[巷弄]/g;
    let m;
    while ((m = re.exec(address))) { cut = m.index; cutEnd = m.index + m[0].length; }
    if (cut >= 0) return address.slice(0, cutEnd).trim();
    return address.replace(/\d+(之\d+)?號.*$/, "").trim();
  }

  // 最後一道備援：只保留第一個數字之前的路名本身（連巷弄都不要），供 OSM 連該巷弄都沒收錄時使用，
  // 定位精度會落在整條路的隨意一點上，只能用來判斷大概方位/最近捷運站，不能拿來當精確位置。
  function majorRoadOnly(address) {
    return address.replace(/\d.*$/, "").trim();
  }

  // 偵測「A路-B路」這種路口交叉點寫法（不是正式門牌，Nominatim 查不到整串）。
  // 抓出兩個路名後分別查詢座標，再取中點做為交叉口的約略位置。
  // 需含「N段」這種路段後綴（如「東華街二段」「中山北路三段」），不然「XX街二段」結尾是「段」字會比對不到
  const ROAD_SUFFIX = "(?:路|街|大道|巷|弄|道)(?:[一二三四五六七八九十]+段)?";
  function parseIntersection(address) {
    const re = new RegExp(`^(.+?${ROAD_SUFFIX})[\\-－/](.+?${ROAD_SUFFIX})$`);
    const m = address.match(re);
    return m ? [m[1].trim(), m[2].trim()] : null;
  }

  // 一條路在 OSM 常被切成好幾段各自獨立的 way，所以要多抓幾筆(limit=6)、把每段的完整線型都攤成座標點，
  // 之後才找得到「這條路真正靠近另一條路的那一段」，而不是隨便抓到不相關的路段
  async function queryRoadPoints(text) {
    const query = encodeURIComponent(text);
    const url = `https://nominatim.openstreetmap.org/search?format=json&polygon_geojson=1&limit=6&countrycodes=tw&q=${query}`;
    const res = await fetch(url, { headers: { Accept: "application/json" } });
    const data = await res.json();
    const points = [];
    (data || []).forEach((item) => {
      if (item.class !== "highway" || !item.geojson) return;
      const geom = item.geojson;
      const lines = geom.type === "LineString" ? [geom.coordinates]
        : geom.type === "MultiLineString" ? geom.coordinates
        : geom.type === "Point" ? [[geom.coordinates]]
        : [];
      lines.forEach((coords) => coords.forEach(([lng, lat]) => points.push({ lat, lng })));
    });
    return points;
  }

  const INTERSECTION_MAX_GAP = 500; // 公尺；兩條路最近的點都還差這麼遠，代表其中一條路名稱有問題(例如少寫路段)，不要硬猜

  // 分別抓兩條路所有路段的座標點，找出兩條路彼此距離最近的一對點，取中點當作路口位置——
  // 而不是各自隨便抓一個代表點就取中點 (那樣常常會抓到兩條路完全不相關、離得很遠的路段)
  async function geocodeIntersection(roadA, roadB, district) {
    const ptsA = await queryRoadPoints(`台北市${district || ""}${roadA}`);
    await sleep(1100); // 尊重 Nominatim 使用限制 (每秒最多1次)
    const ptsB = await queryRoadPoints(`台北市${district || ""}${roadB}`);
    if (!ptsA.length || !ptsB.length) return null;

    let best = null, bestDist = Infinity;
    for (const a of ptsA) {
      for (const b of ptsB) {
        const d = haversine(a.lat, a.lng, b.lat, b.lng);
        if (d < bestDist) { bestDist = d; best = { lat: (a.lat + b.lat) / 2, lng: (a.lng + b.lng) / 2 }; }
      }
    }
    // 兩條路最近的點都差這麼遠，代表其中一條路的名稱可能不完整(例如漏了「N段」)導致抓錯路段，
    // 與其硬給一個可能差了一兩公里的錯誤位置，不如老實回報查無資料，讓使用者自己去 Google 地圖確認
    if (bestDist > INTERSECTION_MAX_GAP) return null;
    return best;
  }

  // 派工系統原始地址有時已包含郵遞區號/台灣/臺北市/行政區/里名，甚至重複兩次（例如「112台灣臺北市臺北市北投區
  // 東華里東華街二段418號」），若不先去掉，之後自動補上的「台北市+行政區」前綴會跟這裡重複，反而讓 Nominatim 查不到。
  // 用迴圈重複剝除，直到剝不動為止，才能處理「臺北市臺北市」這種重複兩次的情況。
  function stripDuplicatePrefix(address, district) {
    let a = address.replace(/^\d{2,6}/, "").trim();
    for (let i = 0; i < 6; i++) {
      const before = a;
      a = a.replace(/^(台灣|臺灣)/, "").trim();
      a = a.replace(/^(台北市|臺北市)/, "").trim();
      if (district) a = a.replace(new RegExp("^" + district), "").trim();
      a = a.replace(/^[一-鿿]{2,4}(里|村)/, "").trim(); // 開頭的里/村名
      if (a === before) break;
    }
    return a || address;
  }

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  function haversine(lat1, lon1, lat2, lon2) {
    const R = 6371000;
    const toRad = (d) => (d * Math.PI) / 180;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  }

  function cacheKeyFor(c) {
    return (c.district || "") + "|" + toHalfWidth(c.address).trim();
  }

  // ---------- 資料同步 (與主視窗共用同一組 localStorage) ----------
  function loadFromStorage() {
    try { cases = JSON.parse(localStorage.getItem(LS_CASES) || "[]"); } catch (e) { cases = []; }
    try { geocache = JSON.parse(localStorage.getItem(LS_GEOCACHE) || "{}"); } catch (e) { geocache = {}; }
    try { rejected = new Set(JSON.parse(localStorage.getItem(LS_REJECTED) || "[]")); } catch (e) { rejected = new Set(); }
    try { marked = new Set(JSON.parse(localStorage.getItem(LS_MARKED) || "[]")); } catch (e) { marked = new Set(); }
    try { lastMonthCases = JSON.parse(localStorage.getItem(LS_LASTMONTH_CASES) || "[]"); } catch (e) { lastMonthCases = []; }
  }

  // 資訊卡的 ❤️／🚫 直接在地圖上改，不用切回主視窗——跟主視窗的 toggleMarked()/toggleRejected() 是
  // 各自獨立的兩份程式碼（不同視窗、不能互相呼叫函式），但寫的是同一組 localStorage key，
  // 主視窗會透過 'storage' 事件收到通知並重畫（見 app.js 的監聽器）
  function toggleMarkedFromMap(id) {
    if (marked.has(id)) marked.delete(id); else marked.add(id);
    localStorage.setItem(LS_MARKED, JSON.stringify(Array.from(marked)));
    plotFromCache(false); // 重畫才能更新資訊卡裡的 ❤️/🤍 圖示；不重新 fitBounds
  }
  function toggleRejectedFromMap(id) {
    if (rejected.has(id)) rejected.delete(id); else rejected.add(id);
    localStorage.setItem(LS_REJECTED, JSON.stringify(Array.from(rejected)));
    plotFromCache(false); // 否決狀態影響點位淡化樣式跟資訊卡文字，都要重畫
  }

  function persistGeocache() {
    localStorage.setItem(LS_GEOCACHE, JSON.stringify(geocache));
  }

  function loadGeocodeSettings() {
    try { return JSON.parse(localStorage.getItem(LS_SETTINGS) || "{}"); } catch (e) { return {}; }
  }

  function activeProvider(settings) {
    const s = settings || loadGeocodeSettings();
    if (s.geocodeProvider === "osm") return "osm";
    return s.googleApiKey ? "google" : "osm"; // "auto" 跟 "google" 沒填金鑰時都退回 osm
  }

  function renderProviderHint() {
    const settings = loadGeocodeSettings();
    const geoProvider = activeProvider(settings);
    const tileProvider = currentTileProvider(settings);
    const el = document.getElementById("providerHint");
    if (!el) return;
    const geoText = geoProvider === "google" ? "🔵 Google 地理編碼" : "⚪ OpenStreetMap 地理編碼";
    const tileText = tileProvider === "google" ? "🗺️ Google 底圖" : "⚪ OpenStreetMap 底圖";
    el.textContent = `${geoText}・${tileText}`;
  }

  // 全部匯入案件都要出現在地圖上，不需要勾選/標記；♡ 標記只用來排路線，不影響地圖顯示範圍
  // 「上月抽查管理」的合併清單一起併進來：地理編碼排隊、圖例、統計都自動一起處理上月案件，
  // 不用另外寫一套平行邏輯；畫點位樣式時用 c.id 是否有「抽」前綴（見 app.js getLastMonthCases()）
  // 判斷要不要用上月專屬的邊框樣式，這樣只要改這一個函式，其他呼叫 allCases() 的地方都自動涵蓋到
  function allCases() {
    return cases.concat(lastMonthCases);
  }

  // ---------- 地圖底圖 (OpenStreetMap 或 Google，依主頁設定) ----------
  let baseLayer = null;
  const loadedScripts = new Set();

  function loadScriptOnce(src) {
    if (loadedScripts.has(src)) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = () => { loadedScripts.add(src); resolve(); };
      s.onerror = () => reject(new Error("載入失敗：" + src));
      document.head.appendChild(s);
    });
  }

  function osmTileLayer() {
    return L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors",
      maxZoom: 19,
    });
  }

  function currentTileProvider(settings) {
    const s = settings || loadGeocodeSettings();
    return s.tileProvider === "google" && s.googleApiKey ? "google" : "osm";
  }

  // 動態載入 Google Maps JavaScript API + leaflet.gridlayer.googlemutant 外掛，
  // 讓 Leaflet 可以顯示 Google 的底圖磚（而不是直接打 Google 的圖磚伺服器，避免違反其使用條款）
  async function ensureGoogleTileSupport(apiKey) {
    if (!(window.google && window.google.maps)) {
      await loadScriptOnce(`https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}`);
    }
    if (!(L.gridLayer && L.gridLayer.googleMutant)) {
      await loadScriptOnce("https://unpkg.com/leaflet.gridlayer.googlemutant@latest/dist/Leaflet.GoogleMutant.js");
    }
  }

  async function setBaseLayer() {
    const settings = loadGeocodeSettings();
    const provider = currentTileProvider(settings);
    let next = null;
    if (provider === "google") {
      try {
        await ensureGoogleTileSupport(settings.googleApiKey);
        next = L.gridLayer.googleMutant({ type: "roadmap" });
      } catch (e) {
        console.error("Google 底圖載入失敗，改用 OpenStreetMap", e);
        next = osmTileLayer();
      }
    } else {
      next = osmTileLayer();
    }
    if (baseLayer) map.removeLayer(baseLayer);
    baseLayer = next.addTo(map);
  }

  // ---------- 地圖 ----------
  async function initMap() {
    map = L.map("map").setView([25.11, 121.52], 12);
    await setBaseLayer();

    mrtLayer = L.layerGroup().addTo(map);
    (window.MRT_STATIONS || []).forEach((st) => {
      L.circleMarker([st.lat, st.lng], {
        radius: 4, color: "#555", weight: 1, fillColor: "#888", fillOpacity: 0.8,
      }).bindTooltip(st.name, { direction: "top" }).addTo(mrtLayer);
    });

    caseLayer = L.layerGroup().addTo(map);
    routeLayer = L.layerGroup().addTo(map);

    // 地圖平移/縮放時，資訊卡開著的話箭頭線要跟著重新指向點位目前的螢幕位置
    map.on("move zoom", updateLeaderLine);
    document.getElementById("markerInfoClose").addEventListener("click", hideMarkerInfo);
  }

  // 用免費的 OSRM 步行路線服務 (OpenStreetMap 資料，無需金鑰) 查詢貼合實際道路的步行路徑；
  // 支援一次帶入多個中途點，回傳整趟路線的完整線型座標 (單位: [lat,lng])
  async function fetchWalkingPath(latlngs) {
    const coordStr = latlngs.map(([lat, lng]) => `${lng},${lat}`).join(";");
    const url = `https://routing.openstreetmap.de/routed-foot/route/v1/foot/${coordStr}?overview=full&geometries=geojson`;
    const res = await fetch(url);
    const data = await res.json();
    if (data.code !== "Ok" || !data.routes || !data.routes[0]) return null;
    return data.routes[0].geometry.coordinates.map(([lng, lat]) => [lat, lng]);
  }

  let routeDrawToken = 0; // 避免路線重新規劃時，前一次還在查詢中的步行路線結果晚到、蓋掉新結果

  // 畫出主視窗規劃好的路線 (依序連線 + 編號標記)，用目前的地理編碼快取查座標；
  // 先畫直線示意，背景查詢實際步行路線，查到後換成貼合道路的路徑，查不到則維持直線示意
  async function drawRoute() {
    const token = ++routeDrawToken;
    routeLayer.clearLayers();
    let route = null;
    try { route = JSON.parse(localStorage.getItem(LS_ROUTE) || "null"); } catch (e) {}
    if (!route || ((!route.order || !route.order.length) && !route.start)) return;

    // 用 allCases()（本月+上月抽查合併），不能只用 cases——路線裡如果排了上月抽查案件的點（id 帶
    // 「抽」前綴），這裡找不到就會被 if (!c) return 悄悄跳過，編號標記跟連線都會漏掉那個點，
    // 路線圖看起來像是提早結束，實際上是抽查案件那幾站完全沒被畫出來
    const byId = new Map(allCases().map((c) => [c.id, c]));
    const latlngs = [];

    if (route.start) {
      latlngs.push([route.start.lat, route.start.lng]);
      L.marker([route.start.lat, route.start.lng], {
        icon: L.divIcon({
          className: "route-stop-icon",
          html: `<div class="route-start-badge">🚇</div>`,
          iconSize: [26, 26],
          iconAnchor: [13, 13],
        }),
      }).bindTooltip(`起點：${route.start.name}`, { direction: "top" }).addTo(routeLayer);
    }

    (route.order || []).forEach((id, idx) => {
      const c = byId.get(id);
      if (!c) return;
      const loc = geocache[cacheKeyFor(c)];
      if (!loc) return;
      latlngs.push([loc.lat, loc.lng]);
      // 這個編號標記畫在跟原本案件圓點同一個座標上面、疊在上層，會直接擋住底下圓點的點擊，
      // 案件排進路線之後點了都沒反應就是這個原因——原本沒有綁點擊事件，這裡補上，
      // 直接沿用底下那個案件圓點（markersById 存的）已經準備好的資訊卡內容，不用重複組一次
      L.marker([loc.lat, loc.lng], {
        icon: L.divIcon({
          className: "route-stop-icon",
          html: `<div class="route-stop-badge">${idx + 1}</div>`,
          iconSize: [22, 22],
          iconAnchor: [11, 11],
        }),
      }).on("click", () => {
        notifyFocusFromMap(id);
        const underlyingMarker = markersById.get(id);
        if (underlyingMarker) showMarkerInfo(underlyingMarker);
      }).addTo(routeLayer);
    });

    if (latlngs.length < 2) return;

    const straightLine = L.polyline(latlngs, { color: "#e0672a", weight: 3, opacity: 0.85, dashArray: "8,6" }).addTo(routeLayer);

    try {
      const walked = await fetchWalkingPath(latlngs);
      if (token !== routeDrawToken) return; // 這次查詢結果已過期 (畫面上又規劃了新路線)，不要套用
      if (walked && walked.length > 1) {
        routeLayer.removeLayer(straightLine);
        L.polyline(walked, { color: "#e0672a", weight: 4, opacity: 0.85 }).addTo(routeLayer);
      }
    } catch (e) {
      console.error("步行路線查詢失敗，維持直線示意", e); // 保留直線示意，仍可看出大致順序與方向
    }
  }

  function renderMapLegend() {
    const cats = Array.from(new Set(allCases().map((c) => c.category)));
    const el = document.getElementById("mapLegend");
    if (!cats.length) { el.innerHTML = `<div class="item"><span class="dot" style="background:#888"></span>灰點 = 捷運站</div>`; return; }
    el.innerHTML = `<div class="item"><span class="dot" style="background:#888"></span>捷運站</div>` +
      cats.map((c) => `<div class="item"><span class="dot" style="background:${colorFor(c)}"></span>${escapeHtml(c)}</div>`).join("");
  }

  function renderSummary() {
    const list = allCases();
    document.getElementById("mapSummaryLine").textContent = !list.length
      ? "尚未匯入任何案件，請回主視窗匯入 xlsx 案件檔"
      : `共 ${list.length} 件匯入案件（與主視窗即時同步）` + (lastMonthCases.length ? `，含 ${lastMonthCases.length} 筆上月抽查` : "");
  }

  // 有人點了地圖標記，通知主視窗跳到清單對應那一列
  function notifyFocusFromMap(id) {
    localStorage.setItem(LS_FOCUS, JSON.stringify({ id, ts: Date.now() }));
  }

  // 主視窗有案件被點擊/勾選分群項目時，飛到該標記並顯示資訊卡
  function handleFocusRequest() {
    let req = null;
    try { req = JSON.parse(localStorage.getItem(LS_FOCUS) || "null"); } catch (e) {}
    if (!req || !req.id || Date.now() - req.ts > 15000) return;
    const marker = markersById.get(req.id);
    if (!marker) return;
    map.flyTo(marker.getLatLng(), 18, { animate: true, duration: 0.6 });
    showMarkerInfo(marker);
  }

  // ---------- 案件標記的資訊卡（取代 Leaflet 預設彈出視窗）----------
  // 原本用 Leaflet 內建的 bindPopup，彈出視窗會直接蓋在點位正上方，把附近的地圖內容擋住。
  // 改成資訊卡固定顯示在地圖右上角，另外畫一條帶箭頭的線從資訊卡指回實際點位，
  // 點位周邊的地圖內容就不會被卡住的視窗擋住了。
  let activeInfoMarker = null;
  function showMarkerInfo(marker) {
    activeInfoMarker = marker;
    const content = document.getElementById("markerInfoContent");
    content.innerHTML = marker._infoHtml || "";
    document.getElementById("markerInfoPanel").classList.add("visible");
    document.getElementById("leaderLineSvg").classList.add("visible");
    updateLeaderLine();
    // innerHTML 整段換掉了，舊的按鈕監聽器（如果有）已經失效，每次顯示資訊卡都要重新綁一次
    const heartBtn = content.querySelector(".heart-btn");
    if (heartBtn) heartBtn.addEventListener("click", (e) => { e.stopPropagation(); toggleMarkedFromMap(heartBtn.dataset.caseId); });
    const rejectBtn = content.querySelector(".reject-btn");
    if (rejectBtn) rejectBtn.addEventListener("click", (e) => { e.stopPropagation(); toggleRejectedFromMap(rejectBtn.dataset.caseId); });
  }
  function hideMarkerInfo() {
    activeInfoMarker = null;
    document.getElementById("markerInfoPanel").classList.remove("visible");
    document.getElementById("leaderLineSvg").classList.remove("visible");
  }
  // 地圖平移/縮放時，讓箭頭線跟著重新指向點位目前的螢幕位置；資訊卡本身固定在右上角不用動。
  // 線的起點原本固定寫死在「資訊卡左下角附近」，但點位可能落在資訊卡的任何方向，固定起點常常
  // 變成線從卡片內容中間穿出去、不太好看。改成跟一般地圖/對話框的指示線做法一樣：算出資訊卡中心
  // 到點位的方向，找這條方向線跟資訊卡外框相交的那一點，線永遠貼著卡片邊框長出來，
  // 不會固定卡在某個角落、也不會穿過卡片裡的文字
  function updateLeaderLine() {
    if (!activeInfoMarker) return;
    const panel = document.getElementById("markerInfoPanel");
    const wrap = document.getElementById("mapWrap");
    const wrapRect = wrap.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    const point = map.latLngToContainerPoint(activeInfoMarker.getLatLng());

    const left = panelRect.left - wrapRect.left;
    const top = panelRect.top - wrapRect.top;
    const right = panelRect.right - wrapRect.left;
    const bottom = panelRect.bottom - wrapRect.top;
    const cx = (left + right) / 2;
    const cy = (top + bottom) / 2;
    const hw = (right - left) / 2;
    const hh = (bottom - top) / 2;
    let dx = point.x - cx;
    let dy = point.y - cy;
    if (dx === 0 && dy === 0) dx = 1; // 退化情況（點位剛好跟卡片中心重疊），避免除以零
    const tx = dx !== 0 ? hw / Math.abs(dx) : Infinity;
    const ty = dy !== 0 ? hh / Math.abs(dy) : Infinity;
    const t = Math.min(tx, ty);
    const x1 = cx + dx * t;
    const y1 = cy + dy * t;

    const line = document.getElementById("leaderLine");
    line.setAttribute("x1", x1);
    line.setAttribute("y1", y1);
    line.setAttribute("x2", point.x);
    line.setAttribute("y2", point.y);
  }

  // 只用目前快取中已有的座標畫圖，不打任何網路請求 (供資料同步/開窗當下立即顯示既有結果)
  // fitView=false：只是重畫既有點位的樣式（例如切換 🚫 否決），座標本身沒有變動，
  // 不需要也不應該連帶重新 fitBounds（否則使用者只是點個 🚫，地圖卻無預警整個 zoom out）
  function plotFromCache(fitView = true) {
    caseLayer.clearLayers();
    markersById.clear();
    const bounds = [];
    const rejectedMarkers = [];
    let cachedCount = 0;

    allCases().forEach((c) => {
      const loc = geocache[cacheKeyFor(c)];
      if (!loc) return;
      cachedCount++;
      const color = colorFor(c.category);
      // 主視窗標記「暫時否決」的案件，這裡同步淡化顯示（保留原本類別顏色跟位置，只降低不透明度、
      // 邊框改灰色），跟主視窗清單「背景白、文字淺灰」是同一個「淡化」語意，方便使用者直接看圖判斷，
      // 不用切回主視窗核對哪些已經否決
      const isRejected = rejected.has(c.id);
      // 「上月抽查管理」的案件 id 全部帶「抽」前綴（app.js getLastMonthCases()），用這個判斷要不要
      // 套用上月專屬的邊框樣式（深藍粗虛線），跟本月案件的白色實框、否決的灰色細框都明顯不同；
      // 否決狀態視覺優先權比較高（本來就代表要淡化不管它是哪個月的），蓋掉上月的邊框樣式
      const isLastMonth = c.id.indexOf("抽") === 0;
      const approxNote = loc.intersection
        ? "<br><i style=\"color:#c9821a\">⚠ 路口交叉點估算位置，僅供大致參考</i>"
        : loc.approx ? "<br><i style=\"color:#c9821a\">⚠ 僅約略定位到路段，非精確門牌</i>" : "";
      const rejectedNote = isRejected ? "<br><b style=\"color:#999\">🚫 已標記暫時否決</b>" : "";
      const lastMonthNote = isLastMonth ? "<br><b style=\"color:#14538f\">📋 上月抽查</b>" : "";
      const marker = L.circleMarker([loc.lat, loc.lng], {
        radius: 7,
        color: isRejected ? "#bbb" : (isLastMonth ? "#14538f" : "#fff"),
        weight: isRejected ? 1.5 : (isLastMonth ? 3 : 1.5),
        fillColor: color,
        fillOpacity: isRejected ? 0.2 : (loc.approx ? 0.8 : 0.95),
        dashArray: isRejected ? null : (isLastMonth ? "5,3" : (loc.approx ? "2,2" : null)),
      }).addTo(caseLayer);
      marker._caseId = c.id;
      const isMarked = marked.has(c.id);
      marker._infoHtml = `<b>${escapeHtml(c.category)}</b><br>案件編號：${escapeHtml(c.id)}<br>${escapeHtml(c.address)}<br>${escapeHtml(c.content).slice(0, 60)}${approxNote}${rejectedNote}${lastMonthNote}
        <div class="marker-info-actions">
          <span class="heart-btn${isMarked ? " marked" : ""}" data-case-id="${escapeHtml(c.id)}" title="標記／取消標記要排路線">${isMarked ? "❤️" : "🤍"}</span>
          <span class="reject-btn${isRejected ? " rejected" : ""}" data-case-id="${escapeHtml(c.id)}" title="暫時否決／取消否決">🚫</span>
        </div>`;
      marker.on("click", () => { notifyFocusFromMap(c.id); showMarkerInfo(marker); });
      markersById.set(c.id, marker);
      bounds.push([loc.lat, loc.lng]);
      if (isRejected) rejectedMarkers.push(marker);
    });

    // 同地址（同座標）的多筆案件會疊在同一個點上，後畫的蓋住先畫的；否決狀態的淡化樣式
    // 一定要疊在最上層，不然使用者剛按完 🚫、地圖上卻還是看到底下另一筆案件的正常顏色，
    // 看起來像沒生效
    rejectedMarkers.forEach((m) => m.bringToFront());

    // 自動定位期間 plotFromCache() 會重複執行、清空重建全部標記；如果資訊卡當時開著，
    // 舊的標記物件已經被丟掉了，要接到新建立的那個上面，資訊卡才不會跟著跑掉或指錯地方
    if (activeInfoMarker) {
      const newMarker = markersById.get(activeInfoMarker._caseId);
      if (newMarker) showMarkerInfo(newMarker);
      else hideMarkerInfo();
    }

    if (fitView && bounds.length) map.fitBounds(bounds, { padding: [30, 30] });
    renderMapLegend();

    const total = allCases().length;
    const missing = total - cachedCount;
    const progress = document.getElementById("geocodeProgress");
    const exportBtn = document.getElementById("btnExportUnlocated");
    if (!total) { progress.textContent = ""; }
    else progress.textContent = `已定位 ${cachedCount}/${total} 筆` + (missing ? `，尚有 ${missing} 筆未定位（按上方按鈕定位）` : "");
    if (exportBtn) exportBtn.style.display = missing ? "inline-block" : "none";
    renderUnlocatedList();
  }

  function unlocatedCases() {
    return allCases().filter((c) => !geocache[cacheKeyFor(c)]);
  }

  function renderUnlocatedList() {
    const el = document.getElementById("unlocatedPanel");
    const list = unlocatedCases();
    if (!list.length) { el.innerHTML = ""; return; }
    el.innerHTML = `
      <h3 style="font-size:13px;color:var(--muted);margin:10px 0 6px">⚠ 未定位案件清單（${list.length} 件，點案件可跳回主視窗清單確認地址）</h3>
      <div style="border:1px solid var(--border);border-radius:8px;background:#fbfcfe;padding:4px 10px;max-height:260px;overflow:auto">
        ${list.map((c) => `
          <div class="station-item" data-case-id="${escapeHtml(c.id)}">
            <span class="badge" style="background:${colorFor(c.category)}22;color:${colorFor(c.category)}">${escapeHtml(c.category)}</span>
            <span class="case-id">#${escapeHtml(c.id)}</span>
            ${escapeHtml(c.address)}
          </div>`).join("")}
      </div>`;
    el.querySelectorAll(".station-item").forEach((row) => {
      row.addEventListener("click", () => notifyFocusFromMap(row.dataset.caseId));
    });
  }

  // ---------- 未定位清單 CSV 匯出 ----------
  function exportUnlocated() {
    const list = unlocatedCases();
    if (!list.length) { alert("目前案件都已定位成功，沒有未定位清單可匯出。"); return; }
    const header = ["案件編號", "案件項目", "行政區", "發生地址", "案件內容"];
    const rows = list.map((c) => [c.id, c.category, c.district, "台北市" + c.address, c.content]);
    const csv = [header, ...rows].map((r) => r.map((v) => `"${String(v || "").replace(/"/g, '""')}"`).join(",")).join("\r\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const ts = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `未定位案件_${ts}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  async function queryNominatim(text) {
    const query = encodeURIComponent(text);
    const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=tw&q=${query}`;
    const res = await fetch(url, { headers: { Accept: "application/json" } });
    const data = await res.json();
    return data && data[0] ? { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) } : null;
  }

  async function queryGoogle(text, apiKey) {
    const query = encodeURIComponent(text);
    const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${query}&region=tw&key=${apiKey}`;
    const res = await fetch(url);
    const data = await res.json();
    const loc = data && data.status === "OK" && data.results[0] && data.results[0].geometry && data.results[0].geometry.location;
    return loc ? { lat: loc.lat, lng: loc.lng } : null;
  }

  // 依「完整地址→cleaned後地址→只到路名→只到主要道路」順序依序嘗試查詢，找到第一筆結果就回傳；
  // 給 Nominatim 用時要傳 delayMs=1100 尊重每秒最多1次的限制，Google 沒有這個限制可以傳 0
  async function runGeocodeAttempts(cleaned, normalized, district, queryFn, delayMs) {
    const attempts = [`台北市${district || ""}${cleaned}`];
    if (cleaned !== normalized) attempts.push(`台北市${district || ""}${normalized}`);
    if (district) attempts.push(`台北市${cleaned}`);
    const road = roadOnly(cleaned);
    if (road && road !== cleaned) attempts.push(`approx:台北市${district || ""}${road}`);
    const majorRoad = majorRoadOnly(cleaned);
    if (majorRoad && majorRoad !== road && majorRoad !== cleaned) attempts.push(`approx:台北市${district || ""}${majorRoad}`);

    let loc = null, approx = false;
    for (let i = 0; i < attempts.length && !loc; i++) {
      if (i > 0 && delayMs) await sleep(delayMs);
      const attempt = attempts[i];
      const isApprox = attempt.startsWith("approx:");
      loc = await queryFn(isApprox ? attempt.slice(7) : attempt);
      if (loc) approx = isApprox;
    }
    return { loc, approx };
  }

  async function geocodeAddress(address, district) {
    const normalized = toHalfWidth(address).trim();
    const cacheKey = (district || "") + "|" + normalized;
    const settings = loadGeocodeSettings();
    const provider = activeProvider(settings);
    const previous = geocache[cacheKey];
    if (!needsGeocode(previous, provider)) return previous;

    const cleaned = stripDuplicatePrefix(normalized, district);
    let loc = null, approx = false, intersection = false;
    const pair = parseIntersection(cleaned) || parseIntersection(normalized);

    if (provider === "google") {
      const queryFn = (text) => queryGoogle(text, settings.googleApiKey);
      // Google 對地址的模糊比對能力較好，交叉口格式直接把兩條路名一起丟給它查詢即可，
      // 不需要像 Nominatim 那樣分別查兩條路的線型再取最近點的中點
      if (pair) {
        loc = await queryFn(`台北市${district || ""}${pair[0]} ${pair[1]}`);
        if (loc) { approx = true; intersection = true; }
      }
      if (!loc) {
        const r = await runGeocodeAttempts(cleaned, normalized, district, queryFn, 0);
        loc = r.loc; approx = r.approx;
      }
    } else {
      // 路口交叉點格式優先用「兩條路分別查、取中點」處理，正常的門牌地址不會被誤判（要求兩段都以路/街/巷/弄結尾）
      if (pair) {
        const mid = await geocodeIntersection(pair[0], pair[1], district);
        if (mid) { loc = mid; approx = true; intersection = true; }
        await sleep(1100); // 尊重 Nominatim 使用限制 (每秒最多1次)
      }
      if (!loc) {
        const r = await runGeocodeAttempts(cleaned, normalized, district, queryNominatim, 1100);
        loc = r.loc; approx = r.approx;
      }
    }

    // 換了服務(例如新加 Google 金鑰)重試卻查無資料時，保留原本已有的約略結果，不要讓案件從地圖上消失
    const result = loc ? { ...loc, approx, intersection, provider } : (previous || null);
    geocache[cacheKey] = result;
    persistGeocache();
    return result;
  }

  // 判斷這筆快取結果是否值得(重)查：從沒查過、先前查無資料，或者「先前只查到約略位置，且當初用的服務跟現在不同」
  // (代表現在可能換了更準的服務，例如新加了 Google 金鑰，值得再試一次看能不能查到精確門牌)
  function needsGeocode(cached, provider) {
    if (!cached) return true;
    if (cached.approx && cached.provider !== provider) return true;
    return false;
  }

  // opts.silent：自動觸發時用（開啟地圖視窗、或主視窗匯入/編輯案件時），不用跳 alert 打斷，
  // 也不需要每次都重新查——只會處理真正還沒定位過的案件
  async function geocodeAllCases(opts) {
    const silent = opts && opts.silent;
    const btn = document.getElementById("btnGeocodeSelected");
    if (btn.disabled) return; // 已經在跑（不管是手動按的還是自動觸發的），避免重疊
    const list = allCases();
    if (!list.length) { if (!silent) alert("尚未匯入任何案件，請先回主視窗匯入 xlsx 案件檔。"); return; }
    const provider = activeProvider();
    // 上月抽查沒比對到上月案件清單明細的（matched:false），地址是空字串，查了也不會有結果，
    // 跳過避免浪費一次查詢等待時間、也避免它一直卡在「未定位」清單裡誤導使用者
    const pending = list.filter((c) => c.address && needsGeocode(geocache[cacheKeyFor(c)], provider));
    if (!pending.length) return;
    btn.disabled = true;
    const progress = document.getElementById("geocodeProgress");
    // plotFromCache() 會重畫全部案件的標記、未定位清單、地圖範圍，案件一多這個重畫本身就不便宜，
    // 每查完一筆就整個重畫一次會變成 O(n²)（實測 136 筆會卡到看起來像當掉）。
    // 改成每查完 5 筆才重畫一次，讓地圖還是會邊查邊補上，但不會每一筆都整套重算
    const REPLOT_EVERY = 5;
    for (let i = 0; i < pending.length; i++) {
      const c = pending[i];
      progress.textContent = `定位中 ${i + 1}/${pending.length}：${c.address}`;
      try { await geocodeAddress(c.address, c.district); } catch (e) {}
      if (provider !== "google") await sleep(1100); // 尊重 Nominatim 使用限制 (每秒最多1次)；Google 沒有這個限制
      if ((i + 1) % REPLOT_EVERY === 0) plotFromCache();
    }
    plotFromCache();
    progress.textContent = "";
    btn.disabled = false;
  }

  // 主視窗變更案件資料時 (localStorage 'storage' 事件只在其他分頁/視窗觸發)，自動同步更新，
  // 並且自動補查新案件/被改過地址的座標，不用使用者手動再按一次「定位案件座標」
  window.addEventListener("storage", (e) => {
    if (e.key === LS_CASES || e.key === LS_GEOCACHE || e.key === LS_LASTMONTH_CASES) {
      loadFromStorage();
      renderSummary();
      plotFromCache();
      drawRoute();
      geocodeAllCases({ silent: true });
    } else if (e.key === LS_REJECTED || e.key === LS_MARKED) {
      loadFromStorage();
      plotFromCache(false); // 只是否決/標記狀態的樣式變化，座標沒變，不要連動 zoom/平移
    } else if (e.key === LS_FOCUS) {
      handleFocusRequest();
    } else if (e.key === LS_ROUTE) {
      drawRoute();
    } else if (e.key === LS_SETTINGS) {
      renderProviderHint();
      setBaseLayer();
    }
  });

  document.addEventListener("DOMContentLoaded", async () => {
    loadFromStorage();
    await initMap();
    document.getElementById("btnGeocodeSelected").addEventListener("click", () => geocodeAllCases());
    document.getElementById("btnExportUnlocated").addEventListener("click", exportUnlocated);
    renderSummary();
    renderProviderHint();
    plotFromCache();
    geocodeAllCases({ silent: true }); // 開啟地圖視窗就自動開始定位，不用再手動按按鈕
    drawRoute();
    handleFocusRequest(); // 若是因為主視窗點案件而剛開啟這個視窗，開窗當下就直接跳過去
  });
})();
