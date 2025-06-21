// ------ 狀態管理與 UI 元件 ------
const state = {
  gapiReady: false, gisReady: false, tokenClient: null, accessToken: null, userProfile: null, driveFolderId: null,
  apiKeys: { google: null, searchEngineId: null, gemini: null, monica: null, driveFolderName: 'MarketingAnalysisTool_Data' },
  isApiReady: false, currentResults: null, initialCompetitorNames: []
};
const ui = {
  authorizeButton: document.getElementById('authorize_button'),
  signOutButton: document.getElementById('signout_button'),
  userProfile: document.getElementById('user-profile'),
  appContainer: document.getElementById('app-container'),
  analysisSection: document.getElementById('analysis-input-section'),
  saveSettingsBtn: document.getElementById('save-settings-btn'),
  startAnalysisBtn: document.getElementById('start-analysis-btn'),
  loadingIndicator: document.getElementById('loading-indicator'),
  spinner: document.querySelector('.spinner'),
  loadingStatus: document.getElementById('loading-status'),
  resultsContainer: document.getElementById('results-tabs-container'),
  resultsActions: document.getElementById('results-actions'),
  historyList: document.getElementById('history-list'),
  competitorSelectPanel: document.getElementById('competitor-select-panel'),
  competitorCheckboxList: document.getElementById('competitor-checkbox-list'),
  selectAllCompetitors: document.getElementById('select-all-competitors'),
  analyzeSelectedBtn: document.getElementById('analyze-selected-btn')
};

window.onload = () => {
  gapi.load('client', () => { state.gapiReady = true; checkAuthReady(); });
  const gisScript = document.createElement('script');
  gisScript.src = 'https://accounts.google.com/gsi/client';
  gisScript.async = true; gisScript.defer = true;
  gisScript.onload = () => {
    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: "733619441496-5gigdpkv42k84no3q5vvh93ths1e8u29.apps.googleusercontent.com",
      scope: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email",
      callback: gisInCallback,
    });
    state.gisReady = true;
    checkAuthReady();
  };
  document.body.appendChild(gisScript);
  setupEventListeners();
};

function setupEventListeners() {
  ui.authorizeButton.onclick = handleAuthClick;
  ui.signOutButton.onclick = handleSignOutClick;
  ui.saveSettingsBtn.onclick = handleSaveSettings;
  ui.startAnalysisBtn.onclick = startInitialAnalysis;
  document.querySelectorAll('.tab-btn').forEach(btn => btn.addEventListener('click', handleTabSwitch));
  document.getElementById('history-list').addEventListener('click', handleHistoryAction);
  document.getElementById('results-actions').addEventListener('click', handleResultAction);
  document.getElementById('results-tabs-container').addEventListener('click', handleTabSwitch);
  ui.selectAllCompetitors.addEventListener('change', function(){
    const all = this.checked;
    document.querySelectorAll('.competitor-checkbox').forEach(cb => cb.checked = all);
  });
  ui.analyzeSelectedBtn.addEventListener('click', analyzeCheckedCompetitors);
}

// ------ Google OAuth 處理 ------
function checkAuthReady() {
  if (state.gapiReady && state.gisReady) {
    gapi.client.init({}).then(() => ui.authorizeButton.classList.remove('hidden'));
  }
}
function gisInCallback(resp) {
  if (resp.error) { showAlert('認證錯誤', resp.error, 'error'); return; }
  state.accessToken = resp;
  gapi.client.setToken(state.accessToken);
  handleSuccessfulLogin();
}
function handleAuthClick() {
  if (gapi.client.getToken() === null) { state.tokenClient.requestAccessToken({ prompt: 'consent' }); }
  else { state.tokenClient.requestAccessToken({ prompt: '' }); }
}
async function handleSuccessfulLogin() {
  ui.authorizeButton.classList.add('hidden');
  ui.userProfile.classList.remove('hidden');
  ui.userProfile.classList.add('flex');
  ui.appContainer.classList.remove('locked');
  const profile = await gapi.client.oauth2.userinfo.get();
  state.userProfile = profile.result;
  document.getElementById('user-avatar').src = state.userProfile.picture;
  document.getElementById('user-name').textContent = state.userProfile.given_name || state.userProfile.name;
  loadSettingsFromLocalStorage();
  if(state.isApiReady) {
    await setupGoogleDrive();
    loadAnalysisHistory();
  }
}
function handleSignOutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revoke(token.access_token, () => {
      gapi.client.setToken('');
      state.accessToken = null;
      state.userProfile = null;
      ui.userProfile.classList.add('hidden');
      ui.authorizeButton.classList.remove('hidden');
      ui.appContainer.classList.add('locked');
      ui.analysisSection.classList.add('locked');
    });
  }
}

// ------ API設定儲存與載入 ------
async function handleSaveSettings() {
  state.apiKeys = {
    driveFolderName: document.getElementById('drive-folder-name').value.trim(),
    google: document.getElementById('google-api-key').value.trim(),
    searchEngineId: document.getElementById('search-engine-id').value.trim(),
    gemini: document.getElementById('ai-gemini-key').value.trim(),
    monica: document.getElementById('ai-monica-key').value.trim()
  };
  if (!state.apiKeys.driveFolderName || !state.apiKeys.google || !state.apiKeys.searchEngineId || (!state.apiKeys.gemini && !state.apiKeys.monica)) {
    showAlert('資訊不完整', '請填寫所有 Google 相關金鑰與至少一組 AI 金鑰及資料夾名稱欄位。', 'warning'); return;
  }
  localStorage.setItem('marketingToolApiKeys', btoa(JSON.stringify(state.apiKeys)));
  state.isApiReady = true;
  ui.analysisSection.classList.remove('locked');
  showAlert('設定已儲存', 'API 已設定完成並啟用，正在設定 Google Drive 資料夾...', 'success');
  await setupGoogleDrive();
  loadAnalysisHistory();
}
function loadSettingsFromLocalStorage() {
  const savedKeys = localStorage.getItem('marketingToolApiKeys');
  if (savedKeys) {
    try {
      const decodedKeys = JSON.parse(atob(savedKeys));
      state.apiKeys = decodedKeys;
      state.isApiReady = true;
      document.getElementById('drive-folder-name').value = state.apiKeys.driveFolderName || 'MarketingAnalysisTool_Data';
      document.getElementById('google-api-key').value = state.apiKeys.google || '';
      document.getElementById('search-engine-id').value = state.apiKeys.searchEngineId || '';
      document.getElementById('ai-gemini-key').value = state.apiKeys.gemini || '';
      document.getElementById('ai-monica-key').value = state.apiKeys.monica || '';
      ui.analysisSection.classList.remove('locked');
    } catch (e) { console.error("無法解析已儲存的 API 金鑰:", e); }
  }
}
async function setupGoogleDrive() {
  if (!state.isApiReady) return;
  try {
    await gapi.client.load('drive', 'v3');
    const folderName = state.apiKeys.driveFolderName;
    const response = await gapi.client.drive.files.list({
      q: `name='${folderName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields: 'files(id)'
    });
    if (response.result.files && response.result.files.length > 0) {
      state.driveFolderId = response.result.files[0].id;
    } else {
      const fileMetadata = { 'name': folderName, 'mimeType': 'application/vnd.google-apps.folder' };
      const createResponse = await gapi.client.drive.files.create({ resource: fileMetadata, fields: 'id' });
      state.driveFolderId = createResponse.result.id;
    }
    showAlert('Google Drive 已就緒', `將使用資料夾 "${folderName}" 存取紀錄。`, 'info');
  } catch(error) {
    console.error("Google Drive 設定失敗:", error);
    showAlert('Drive 設定失敗', `無法存取或建立資料夾: ${error.message}`, 'error');
  }
}

// ------ 條件與Prompt處理 ------
function getSearchContext() {
  const industry = document.getElementById('industry').value.trim();
  const sector = document.getElementById('sector').value.trim();
  const product = document.getElementById('product').value.trim();
  const company = document.getElementById('company').value.trim();
  const businessModel = document.querySelector('input[name="business-model"]:checked').value;
  let topic = '';
  if (industry) topic += `產業:${industry} `;
  if (sector) topic += `行業:${sector} `;
  if (product) topic += `產品:${product} `;
  if (company) topic += `公司:${company}`;
  return { topic: topic.trim(), businessModel, industry, sector, product, company };
}

// ------ 分析主流程 ------
async function startInitialAnalysis() {
  if (!state.isApiReady) { showAlert('API 尚未就緒', '請先儲存您的 API 設定。', 'error'); return; }
  const { industry, sector, product, company } = getSearchContext();
  if (!industry && !sector && !product && !company) { showAlert('需要輸入', '請至少填寫一個分析欄位。', 'warning'); return; }

  ui.loadingIndicator.classList.remove('hidden');
  ui.spinner.classList.remove('hidden');
  ui.loadingStatus.textContent = '步驟 1/2: 正在生成關鍵字與競爭對手名單...';
  ui.resultsContainer.classList.add('hidden');
  ui.resultsActions.classList.add('hidden');
  ui.competitorSelectPanel.classList.add('hidden');
  try {
    const keywordsResponse = await getAiKeywordAnalysis();
    state.currentResults = { keywords: keywordsResponse.keywords || [] };
    renderPartialResults('keywords');
    // 產生競爭對手名稱
    let aiPromises = [];
    if (state.apiKeys.gemini) aiPromises.push(getAiCompetitorNames_Gemini().catch(e => { showAlert('Gemini 失敗', 'Gemini未取得競爭對手'); return []; }));
    if (state.apiKeys.monica) aiPromises.push(getAiCompetitorNames_Monica().catch(e => { showAlert('Monica 失敗', 'Monica未取得競爭對手'); return []; }));
    if (!aiPromises.length) { showAlert('錯誤', '請至少填寫一組AI API金鑰', 'warning'); return; }
    const aiResults = await Promise.all(aiPromises);
    let mergedNames = [];
    if (aiResults[0]) aiResults[0].forEach(n => mergedNames.push({ name: n, source: state.apiKeys.gemini ? 'Gemini' : '' }));
    if (aiResults[1]) aiResults[1].forEach(n => {
      let idx = mergedNames.findIndex(i => i.name === n);
      if (idx === -1) mergedNames.push({ name: n, source: 'Monica' });
      else mergedNames[idx].source = mergedNames[idx].source === 'Gemini' ? '兩者' : 'Monica';
    });
    if (!mergedNames.length) { showAlert('提示', '查無競爭對手，請調整條件。', 'info'); return; }
    state.initialCompetitorNames = mergedNames;
    showCompetitorSelectPanel(mergedNames);
    ui.loadingStatus.textContent = '請在下方勾選競爭對手並按「進行分析」';
    ui.spinner.classList.add('hidden');
  } catch (error) {
    showAlert('分析失敗', '步驟一失敗: ' + error?.message, 'error');
    ui.spinner.classList.add('hidden');
  }
}
async function getAiKeywordAnalysis() {
  const context = getSearchContext();
  const prompt = `
你是一位專精於${context.businessModel}領域的市場分析專家。
主題：${context.topic}
請根據${context.businessModel}模式，生成10個最重要的關鍵字，針對每個關鍵字給出「搜尋量級」(高/中/低)、「類型」(B2B專業、B2C消費、B2B2C平台等)。
請根據下列定義對每個關鍵字分類：
- B2B：專為企業設計的服務或產品
- B2C：直接服務消費者的內容
- B2B2C：同時對企業與消費者雙邊的服務（如外送平台、電商平台，需明確有雙邊端）
請以 JSON 格式回傳，如：{"keywords":[{"keyword":"xxx","volume":"高","type":"B2B專業"},...]}
  `;
  if (state.apiKeys.gemini) return safeJsonParse(await callGeminiAPI(prompt, state.apiKeys.gemini));
  else if (state.apiKeys.monica) return safeJsonParse(await callMonicaAPI(prompt, state.apiKeys.monica));
  else throw new Error('未設定AI金鑰');
}
async function getAiCompetitorNames_Gemini() {
  const context = getSearchContext();
  const prompt = `
你是一位${context.businessModel}領域的市場分析專家。
主題：${context.topic}
請僅列出5~10個最具代表性的競爭對手品牌名稱（只要名稱，不需其他細節），請以JSON格式回傳：
{"competitorNames":["品牌1","品牌2",...]}
  `;
  const resp = safeJsonParse(await callGeminiAPI(prompt,state.apiKeys.gemini));
  return resp.competitorNames || [];
}
async function getAiCompetitorNames_Monica() {
  const context = getSearchContext();
  const prompt = `
你是一位${context.businessModel}行業的競爭對手情報分析師。
主題：${context.topic}
請列出5~10個主要競爭對手品牌名稱（僅名稱即可），以JSON格式回傳，如：
{"competitorNames":["品牌A","品牌B",...]}
  `;
  const resp = safeJsonParse(await callMonicaAPI(prompt,state.apiKeys.monica));
  return resp.competitorNames || [];
}
function showCompetitorSelectPanel(competitorList) {
  const panel = ui.competitorSelectPanel;
  const list = ui.competitorCheckboxList;
  list.innerHTML = competitorList.map((item, idx) =>
    `<label class="flex items-center gap-2 py-2 px-3 rounded-lg hover:bg-techprimary/10 cursor-pointer transition">
      <input type="checkbox" class="competitor-checkbox accent-techaccent scale-125" value="${item.name}" checked>
      <span class="tracking-wide font-semibold text-techaccent">${item.name}</span>
      <span class="ml-2 text-xs text-techprimary/70 bg-techbg px-2 py-0.5 rounded">${item.source}</span>
    </label>`
  ).join('');
  panel.classList.remove('hidden');
  ui.selectAllCompetitors.checked = true;
}
async function analyzeCheckedCompetitors() {
  const selected = Array.from(document.querySelectorAll('.competitor-checkbox'))
    .filter(cb => cb.checked)
    .map(cb => cb.value);
  if (!selected.length) { showAlert("未選擇", "請至少勾選一個競爭對手。", "warning"); return; }
  ui.loadingIndicator.classList.remove('hidden');
  ui.spinner.classList.remove('hidden');
  ui.loadingStatus.textContent = "步驟 2/2: 正在取得競爭對手細部資料...";
  ui.resultsContainer.classList.add('hidden');
  ui.resultsActions.classList.add('hidden');
  try {
    let aiPromises = [];
    if (state.apiKeys.gemini) aiPromises.push(getAiCompetitorDetails_Gemini(selected).catch(e => { showAlert('Gemini細節失敗',e.message,'warning'); return []; }));
    if (state.apiKeys.monica) aiPromises.push(getAiCompetitorDetails_Monica(selected).catch(e => { showAlert('Monica細節失敗',e.message,'warning'); return []; }));
    if (!aiPromises.length) { showAlert('錯誤','請至少填寫一組AI API金鑰','warning'); return; }
    const aiResults = await Promise.all(aiPromises);
    let allDetails = [];
    if (aiResults[0]) aiResults[0].forEach(item => allDetails.push({ ...item, aiSource: state.apiKeys.gemini ? 'Gemini' : '' }));
    if (aiResults[1]) aiResults[1].forEach(item => {
      let existing = allDetails.find(d => d.brandName === item.brandName);
      if (existing) {
        existing.aiSource = existing.aiSource === 'Gemini' ? '兩者' : 'Monica';
        for (const key in item) if (item[key] && !existing[key]) existing[key] = item[key];
      } else {
        allDetails.push({ ...item, aiSource: 'Monica' });
      }
    });
    state.currentResults.competitors = allDetails;
    renderPartialResults('competitors');
    ui.resultsContainer.classList.remove('hidden');
    ui.resultsActions.classList.remove('hidden');
    showAlert('分析完成', '競爭對手詳細資料已取得。', 'success');
  } catch (e) {
    showAlert('分析失敗', 'AI分析競爭對手細節失敗。', 'error');
  } finally {
    ui.spinner.classList.add('hidden');
  }
}
async function getAiCompetitorDetails_Gemini(competitorNames) {
  const context = getSearchContext();
  const prompt = `
請針對以下競爭對手品牌，列出每個品牌的詳細資料：品牌名稱、業務摘要、官方網站、公司類型、LinkedIn公司頁、股票代號、郵遞區號。
競爭對手品牌清單：[${competitorNames.join('，')}]
請以JSON格式回傳，如：
{"competitors":[{"brandName":"xxx","summary":"...","websiteUrl":"...","companyType":"...","linkedinUrl":"...","stockCode":"...","zipcode":"..."}]}
  `;
  const resp = safeJsonParse(await callGeminiAPI(prompt,state.apiKeys.gemini));
  return resp.competitors || [];
}
async function getAiCompetitorDetails_Monica(competitorNames) {
  const context = getSearchContext();
  const prompt = `
請根據主題：${context.topic}，針對競爭對手品牌（${competitorNames.join("，")}），產生每個品牌的詳細資料，包括品牌名稱、業務摘要、網站、公司類型、LinkedIn頁、股票代號、郵遞區號。請以JSON格式回傳：
{"competitors":[{"brandName":"...","summary":"...","websiteUrl":"...","companyType":"...","linkedinUrl":"...","stockCode":"...","zipcode":"..."}]}
  `;
  const resp = safeJsonParse(await callMonicaAPI(prompt,state.apiKeys.monica));
  return resp.competitors || [];
}
async function callGeminiAPI(prompt, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
  });
  if (!response.ok) {
    const errorBody = await response.json();
    throw new Error(`Gemini API Error: ${errorBody?.error?.message || '未知錯誤'}`);
  }
  const data = await response.json();
  if (!data.candidates || !data.candidates[0].content.parts[0].text) {
    throw new Error("Gemini API 回應格式無效。");
  }
  return data.candidates[0].content.parts[0].text;
}
async function callMonicaAPI(prompt, apiKey) {
  const url = `https://api.monica.im/api/v1/chat/completions`;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
    body: JSON.stringify({ model: "monica-3.5", messages: [{ role: "user", content: prompt }] })
  });
  if (!response.ok) {
    const errorBody = await response.json();
    throw new Error(`Monica API Error: ${errorBody?.error?.message || '未知錯誤'}`);
  }
  const data = await response.json();
  if (!data.choices || !data.choices[0].message.content) {
    throw new Error("Monica API 回應格式無效。");
  }
  return data.choices[0].message.content;
}

// ------ 結果渲染 ------
function renderPartialResults(type) {
  ui.resultsContainer.classList.remove('hidden');
  ui.loadingIndicator.classList.add('hidden');
  if (type === 'keywords') {
    const panel = document.getElementById('keywords-result-panel');
    if (!state.currentResults.keywords || !state.currentResults.keywords.length) {
      panel.innerHTML = '<div class="text-gray-500">查無關鍵字資料。</div>';
      return;
    }
    panel.innerHTML = `<table class="w-full text-left text-sm">
      <thead class="bg-gray-900"><tr class="border-b border-techprimary/30"><th class="py-2 px-2">關鍵字</th><th>搜尋量級</th><th>類型</th></tr></thead>
      <tbody>${state.currentResults.keywords.map(k => `<tr class="border-b border-techprimary/20"><td class="py-2 px-2">${k.keyword}</td><td>${k.volume}</td><td>${k.type}</td></tr>`).join('')}</tbody>
    </table>`;
  }
  if (type === 'competitors') {
    const competitorsPanel = document.getElementById('competitors-result-panel');
    if (!state.currentResults.competitors || !state.currentResults.competitors.length) {
      competitorsPanel.innerHTML = '<div class="text-gray-500">查無競爭對手資料。</div>';
    } else {
      competitorsPanel.innerHTML = (state.currentResults.competitors || []).map(c =>
        `<div class="p-4 rounded-xl bg-gradient-to-br from-techbg to-techpanel shadow-lg border border-techprimary/20 mb-4">
          <div class="flex gap-3 items-center">
            <span class="inline-block w-2 h-2 rounded-full bg-techaccent animate-pulse"></span>
            <span class="font-bold text-lg text-techaccent">${c.brandName}</span>
            <span class="text-xs bg-techprimary/20 text-techprimary px-2 py-0.5 rounded ml-2">${c.aiSource || ''}</span>
          </div>
          <p class="mt-2 text-gray-300">${c.summary || ''}</p>
          <a href="${c.websiteUrl}" target="_blank" class="block text-techaccent underline hover:text-techprimary mt-1 transition">${c.websiteUrl||''}</a>
          <div class="mt-2 text-sm text-techprimary/60">
            公司類型：${c.companyType || '-'}<br/>
            股號：${c.stockCode || '-'}<br/>
            郵遞區號：${c.zipcode || '-'}
          </div>
        </div>`
      ).join('');
    }
    // LinkedIn
    const linkedinPanel = document.getElementById('linkedin-result-panel');
    if (!state.currentResults.competitors || !state.currentResults.competitors.length) {
      linkedinPanel.innerHTML = '<div class="text-gray-500">查無LinkedIn資料。</div>';
    } else {
      linkedinPanel.innerHTML = (state.currentResults.competitors || []).map(c =>
        `<div class="p-4 rounded-xl bg-gradient-to-br from-techbg to-techpanel shadow-lg border border-techprimary/20 mb-4">
          <div class="flex gap-3 items-center">
            <span class="inline-block w-2 h-2 rounded-full bg-techaccent animate-pulse"></span>
            <span class="font-bold text-lg text-techaccent">${c.brandName}</span>
            <span class="text-xs bg-techprimary/20 text-techprimary px-2 py-0.5 rounded ml-2">${c.aiSource || ''}</span>
          </div>
          <p class="mt-1 text-gray-200">公司類型：${c.companyType || ''}</p>
          <a href="${c.linkedinUrl}" target="_blank" class="block text-techaccent underline hover:text-techprimary mt-1 transition">${c.linkedinUrl||''}</a>
          <div class="mt-2 text-sm text-techprimary/60">
            股號：${c.stockCode || '-'}<br/>
            郵遞區號：${c.zipcode || '-'}
          </div>
        </div>`
      ).join('');
    }
  }
}

// ------ 匯出與儲存 ------
function exportToExcel() {
  if (!state.currentResults || !state.currentResults.keywords) {
    showAlert("無法匯出", "目前尚未有完整分析結果。", "warning");
    return;
  }
  const wb = XLSX.utils.book_new();
  // 關鍵字
  const keywordsSheet = XLSX.utils.json_to_sheet(state.currentResults.keywords || []);
  XLSX.utils.book_append_sheet(wb, keywordsSheet, "A_關鍵字分析");
  // 競爭對手網站
  const competitorsSheetData = (state.currentResults.competitors || []).map(c => ({
      '品牌名稱': c.brandName,
      '業務摘要': c.summary,
      '網站': c.websiteUrl,
      '公司類型': c.companyType,
      '股票代號': c.stockCode,
      '郵遞區號': c.zipcode,
      'AI來源': c.aiSource
  }));
  const competitorsSheet = XLSX.utils.json_to_sheet(competitorsSheetData);
  XLSX.utils.book_append_sheet(wb, competitorsSheet, "B_競爭對手網站");
  // LinkedIn
  const linkedinSheetData = (state.currentResults.competitors || []).map(c => ({
      '品牌名稱': c.brandName,
      '公司類型': c.companyType,
      'LinkedIn': c.linkedinUrl,
      '股票代號': c.stockCode,
      '郵遞區號': c.zipcode,
      'AI來源': c.aiSource
  }));
  const linkedinSheet = XLSX.utils.json_to_sheet(linkedinSheetData);
  XLSX.utils.book_append_sheet(wb, linkedinSheet, "C_LinkedIn分析");
  const context = getSearchContext();
  XLSX.writeFile(wb, `${context.topic || '分析'}_Analysis.xlsx`);
}
async function saveAnalysisToDrive() {
  if (!state.currentResults || !state.driveFolderId) {
    showAlert('無法儲存', '請先生成報告並確保已設定 Google Drive 資料夾。', 'warning');
    return;
  }
  const context = getSearchContext();
  const metadata = {
    name: `${context.topic}_${new Date().toISOString()}.json`,
    mimeType: 'application/json',
    parents: [state.driveFolderId]
  };
  const boundary = 'foo_bar_baz';
  const multipartRequestBody = 
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
      `--${boundary}\r\nContent-Type: application/json\r\n\r\n${JSON.stringify({results: state.currentResults, queriedAt: new Date().toISOString()})}\r\n--${boundary}--`;
  try {
    await gapi.client.request({
      path: '/upload/drive/v3/files',
      method: 'POST',
      params: { uploadType: 'multipart' },
      headers: { 'Content-Type': 'multipart/related; boundary="' + boundary + '"' },
      body: multipartRequestBody
    });
    showAlert('儲存成功', '分析報告已儲存至您的 Google Drive。', 'success');
    loadAnalysisHistory();
  } catch (e) { showAlert('儲存失敗', '無法將報告儲存至 Google Drive。', 'error'); }
}
async function loadAnalysisHistory() {
  if (!state.driveFolderId) {
    ui.historyList.innerHTML = '<p class="text-gray-500">請先在設定中指定一個 Google Drive 資料夾名稱並儲存。</p>';
    return;
  }
  const response = await gapi.client.drive.files.list({ q: `'${state.driveFolderId}' in parents and trashed=false`, fields: 'files(id, name, createdTime)', orderBy: 'createdTime desc' });
  const files = response.result.files;
  ui.historyList.innerHTML = '';
  if (files && files.length > 0) {
    files.forEach(file => {
      const fileEl = document.createElement('div');
      fileEl.className = 'border p-3 rounded-lg flex justify-between items-center';
      fileEl.innerHTML = `<div><p class="font-semibold">${file.name.replace('.json', '')}</p><p class="text-sm text-gray-500">儲存於: ${new Date(file.createdTime).toLocaleDateString()}</p></div><button class="load-history-btn bg-blue-100 px-3 py-1 rounded text-blue-800 font-semibold" data-id="${file.id}">載入</button>`;
      ui.historyList.appendChild(fileEl);
    });
  } else { ui.historyList.innerHTML = '<p class="text-gray-500">此資料夾中沒有歷史紀錄。</p>'; }
}
async function handleHistoryAction(e) {
  if (e.target.classList.contains('load-history-btn')) {
    const fileId = e.target.dataset.id;
    try {
      const response = await gapi.client.drive.files.get({ fileId: fileId, alt: 'media' });
      state.currentResults = response.result.results;
      renderPartialResults('keywords');
      renderPartialResults('competitors');
      document.querySelector('.tab-btn[data-tab="results"]').click();
      showAlert('紀錄已載入', '成功從歷史紀錄載入分析報告。', 'info');
    } catch (error) { showAlert('載入失敗', '無法載入所選的歷史紀錄。', 'error'); }
  }
}
function handleResultAction(e) {
  if (e.target.id === 'save-analysis-btn') saveAnalysisToDrive();
  else if (e.target.id === 'export-excel-btn') exportToExcel();
}
function handleTabSwitch(e) {
  const target = e.target.closest('.tab-btn');
  if (target) {
    if (target.classList.contains('result-tab-btn')) {
      const tab = target.dataset.resultTab;
      document.querySelectorAll('.result-tab-btn').forEach(b => b.classList.remove('active'));
      target.classList.add('active');
      document.querySelectorAll('.result-panel').forEach(p => p.classList.remove('active'));
      document.getElementById(`${tab}-result-panel`).classList.add('active');
    } else {
      const tab = target.dataset.tab;
      document.querySelectorAll('.tab-btn:not(.result-tab-btn)').forEach(b => b.classList.remove('active'));
      target.classList.add('active');
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
      document.getElementById(`${tab}-panel`).classList.add('active');
    }
  }
}
function showAlert(title, message, type = 'info') {
  const colors = { info: 'bg-blue-100 border-blue-500 text-blue-700', success: 'bg-green-100 border-green-500 text-green-700', warning: 'bg-yellow-100 border-yellow-500 text-yellow-700', error: 'bg-red-100 border-red-500 text-red-700' };
  const alertEl = document.createElement('div');
  alertEl.className = `p-4 mb-4 rounded-md shadow-lg border-l-4 ${colors[type]}`;
  alertEl.innerHTML = `<div class="flex"><div class="ml-3"><p class="font-bold">${title}</p><p class="text-sm">${message}</p></div><button class="ml-auto -mx-1.5 -my-1.5 p-1.5 rounded-lg inline-flex items-center justify-center text-${type === 'error' ? 'red' : 'gray'}-500 hover:bg-gray-200" onclick="this.parentElement.parentElement.remove()"><span class="sr-only">關閉</span>✕</button></div>`;
  document.getElementById('alert-container').prepend(alertEl);
  setTimeout(() => alertEl.remove(), 6000);
}
function safeJsonParse(str) {
  try {
    const cleanStr = str.replace(/```json/g, '').replace(/```/g, '').trim();
    const obj = JSON.parse(cleanStr);
    if (!obj.keywords || !Array.isArray(obj.keywords)) obj.keywords = [];
    if (!obj.competitors || !Array.isArray(obj.competitors)) obj.competitors = [];
    if (!obj.competitorNames || !Array.isArray(obj.competitorNames)) obj.competitorNames = [];
    return obj;
  } catch (e) {
    console.error("JSON 解析失敗:", str);
    throw new Error("AI 回應的資料格式錯誤，無法解析。");
  }
}