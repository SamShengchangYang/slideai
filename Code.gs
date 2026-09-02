/**
 * AI簡報力工作坊 — 互動學習系統（Google Apps Script 全端）
 * 部署方式：以我的身分執行 / 任何人皆可存取
 * 資料庫：本試算表（首次執行 setup() 自動建表與種子資料）
 *
 * RBAC：admin（管理者）/ learner（學員）
 */

var SS = null;
function ss_() { if (!SS) SS = SpreadsheetApp.getActiveSpreadsheet(); return SS; }

// ====================== 初始化（手動執行一次） ======================
function setup() {
  var ss = ss_();
  ensureSheet_('Users',      ['uid','account','passHash','salt','name','role','dept','createdAt']);
  ensureSheet_('Progress',   ['uid','pageId','status','updatedAt']);
  ensureSheet_('Responses',  ['rid','uid','pageId','qid','answer','analysis','feedback','createdAt']);
  ensureSheet_('Pages',      ['pageId','order','title','subtitle','contentJson','questionsJson','active']);
  ensureSheet_('Live',       ['qid','title','options','explain','status','createdAt']);
  ensureSheet_('LiveAnswers',['qid','uid','choice','reason','at']);
  seedPages_();
  seedQuiz_();
  bankSheet_();
  // 預設管理者帳號 admin / 請立即改密碼
  if (findUserByAccount_('admin') == null) createUser_('admin', 'cmuh2026', '系統管理者', 'admin', '教學部');
  return 'setup 完成';
}

function ensureSheet_(name, headers) {
  var ss = ss_(); var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

// ====================== 密碼與 Session ======================
function hash_(pass, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass + '::' + salt);
  return raw.map(function(b){ b = (b < 0 ? b + 256 : b); return ('0'+b.toString(16)).slice(-2); }).join('');
}
function newToken_() { return Utilities.getUuid().replace(/-/g,'') + Date.now().toString(36); }

function cachePutSession_(token, uid, role) {
  CacheService.getScriptCache().put('sess_'+token, JSON.stringify({uid:uid, role:role}), 21600); // 6hr
}
function session_(token) {
  if (!token) return null;
  var v = CacheService.getScriptCache().get('sess_'+token);
  return v ? JSON.parse(v) : null;
}
function requireRole_(token, role) {
  var s = session_(token);
  if (!s) throw new Error('SESSION_EXPIRED');
  if (role && s.role !== role) throw new Error('FORBIDDEN');
  return s;
}

// ====================== 使用者 ======================
function rows_(name) {
  var sh = ss_().getSheetByName(name);
  var data = sh.getDataRange().getValues();
  var head = data.shift();
  return data.map(function(r){ var o={}; head.forEach(function(h,i){o[h]=r[i];}); return o; });
}
function findUserByAccount_(account) {
  return rows_('Users').filter(function(u){ return String(u.account) === String(account); })[0] || null;
}
function createUser_(account, pass, name, role, dept) {
  var salt = Utilities.getUuid().slice(0,8);
  ss_().getSheetByName('Users').appendRow([Utilities.getUuid().slice(0,8), account, hash_(pass, salt), salt, name, role, dept||'', new Date()]);
}

// ====================== 公開 API（google.script.run） ======================

/** 註冊（學員自助）。回傳 {ok} 或 {error} */
function apiRegister(account, pass, name, dept) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    account = String(account||'').trim(); name = String(name||'').trim();
    if (!account || !pass || !name) return {error:'請完整填寫帳號、密碼與姓名'};
    if (findUserByAccount_(account)) return {error:'此帳號已存在，請直接登入'};
    createUser_(account, pass, name, 'learner', dept||'');
    return {ok:true};
  } finally { lock.releaseLock(); }
}

/** 登入。回傳 {token, role, name, uid} */
function apiLogin(account, pass) {
  var u = findUserByAccount_(String(account||'').trim());
  if (!u) return {error:'帳號不存在'};
  if (hash_(pass, u.salt) !== u.passHash) return {error:'密碼錯誤'};
  var token = newToken_();
  cachePutSession_(token, u.uid, u.role);
  return {token:token, role:u.role, name:u.name, uid:u.uid};
}

/** 取得課程頁清單與個人進度（學員/管理者皆可） */
function apiGetPages(token) {
  var s = requireRole_(token, null);
  var pages = rows_('Pages').filter(function(p){ return p.active !== false && p.active !== 'FALSE'; })
    .sort(function(a,b){ return a.order-b.order; })
    .map(function(p){ return {pageId:p.pageId, order:p.order, title:p.title, subtitle:p.subtitle,
      content:JSON.parse(p.contentJson||'[]'), questions:JSON.parse(p.questionsJson||'[]')}; });
  var prog = {};
  rows_('Progress').filter(function(r){ return r.uid===s.uid; }).forEach(function(r){ prog[r.pageId]=r.status; });
  var mine = rows_('Responses').filter(function(r){ return r.uid===s.uid; })
    .map(function(r){ return {pageId:r.pageId, qid:r.qid, answer:r.answer, feedback:r.feedback}; });
  var u = rows_('Users').filter(function(x){ return x.uid===s.uid; })[0] || {};
  return {pages:pages, progress:prog, myResponses:mine, me:{name:u.name, dept:u.dept, role:u.role, account:u.account}};
}

/** 更新進度：status = 'doing' | 'done' */
function apiSetProgress(token, pageId, status) {
  var s = requireRole_(token, null);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = ss_().getSheetByName('Progress');
    var data = sh.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][0]===s.uid && data[i][1]===pageId) { sh.getRange(i+1,3).setValue(status); sh.getRange(i+1,4).setValue(new Date()); return {ok:true}; }
    }
    sh.appendRow([s.uid, pageId, status, new Date()]);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

/** 提交反思回答 → 立即做規則式概念分析並回傳即時回饋 */
function apiSubmitAnswer(token, pageId, qid, answer) {
  var s = requireRole_(token, null);
  var page = rows_('Pages').filter(function(p){ return p.pageId===pageId; })[0];
  if (!page) return {error:'找不到頁面'};
  var q = JSON.parse(page.questionsJson||'[]').filter(function(x){ return x.qid===qid; })[0];
  var analysis = analyzeAnswer_(answer, q);
  var feedback = buildFeedback_(analysis, q);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    ss_().getSheetByName('Responses').appendRow([Utilities.getUuid().slice(0,8), s.uid, pageId, qid,
      String(answer||''), JSON.stringify(analysis), feedback, new Date()]);
  } finally { lock.releaseLock(); }
  return {ok:true, feedback:feedback, analysis:analysis};
}

/** 規則式分析：關鍵概念覆蓋 + 常見誤區偵測 */
function analyzeAnswer_(answer, q) {
  answer = String(answer||'');
  var hit=[], miss=[], traps=[];
  (q && q.concepts || []).forEach(function(c){
    var found = (c.keys||[]).some(function(k){ return answer.indexOf(k) > -1; });
    (found ? hit : miss).push(c.label);
  });
  (q && q.traps || []).forEach(function(t){
    var found = (t.keys||[]).some(function(k){ return answer.indexOf(k) > -1; });
    if (found) traps.push(t.label);
  });
  var score = (q && q.concepts && q.concepts.length) ? Math.round(hit.length / q.concepts.length * 100) : null;
  return {hit:hit, miss:miss, traps:traps, coverage:score, length:answer.length};
}

function buildFeedback_(a, q) {
  var msg = [];
  if (a.hit.length) msg.push('你已掌握：' + a.hit.join('、') + '。');
  if (a.miss.length) msg.push('可以再想想：' + a.miss.join('、') + '——回頭對照本頁內容，這些概念和你的回答有什麼關聯？');
  if (a.traps.length) msg.push('留意常見誤區：' + a.traps.join('、') + '。');
  if (!msg.length) msg.push('已收到你的回答，講師將於課後給你回饋。');
  if (q && q.hint && a.miss.length) msg.push('提示：' + q.hint);
  return msg.join('\n');
}

// ====================== 管理者 API ======================

function apiAdminDashboard(token) {
  requireRole_(token, 'admin');
  var users = rows_('Users');
  var prog = rows_('Progress');
  var resp = rows_('Responses');
  var pages = rows_('Pages').sort(function(a,b){ return a.order-b.order; });

  var board = users.filter(function(u){ return u.role==='learner'; }).map(function(u){
    var p = {}; prog.filter(function(r){ return r.uid===u.uid; }).forEach(function(r){ p[r.pageId]=r.status; });
    return {uid:u.uid, name:u.name, dept:u.dept, account:u.account, progress:p,
      answers: resp.filter(function(r){ return r.uid===u.uid; }).length};
  });

  // 誤區/概念缺口彙總
  var missCount = {}, trapCount = {};
  resp.forEach(function(r){
    try {
      var a = JSON.parse(r.analysis||'{}');
      (a.miss||[]).forEach(function(m){ missCount[m]=(missCount[m]||0)+1; });
      (a.traps||[]).forEach(function(t){ trapCount[t]=(trapCount[t]||0)+1; });
    } catch(e){}
  });

  var answers = resp.map(function(r){
    var u = users.filter(function(x){ return x.uid===r.uid; })[0] || {};
    return {rid:r.rid, name:u.name||r.uid, pageId:r.pageId, qid:r.qid, answer:r.answer,
      analysis:r.analysis, feedback:r.feedback,
      at: r.createdAt ? Utilities.formatDate(new Date(r.createdAt), 'Asia/Taipei', 'MM/dd HH:mm') : ''};
  }).reverse().slice(0,300);

  return {board:board, missCount:missCount, trapCount:trapCount, answers:answers,
    pageMeta: pages.map(function(p){ return {pageId:p.pageId, title:p.title, active:String(p.active)!=='FALSE'}; })};
}

/** 帳號管理：list / create / resetPass / setRole / remove */
function apiAdminUsers(token, action, payload) {
  requireRole_(token, 'admin');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = ss_().getSheetByName('Users');
    if (action==='list') return {users: rows_('Users').map(function(u){ return {uid:u.uid,account:u.account,name:u.name,role:u.role,dept:u.dept}; })};
    if (action==='create') {
      if (findUserByAccount_(payload.account)) return {error:'帳號已存在'};
      createUser_(payload.account, payload.pass, payload.name, payload.role||'learner', payload.dept||'');
      return {ok:true};
    }
    var data = sh.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][0]===payload.uid) {
        if (action==='resetPass') { var salt=Utilities.getUuid().slice(0,8); sh.getRange(i+1,3).setValue(hash_(payload.pass,salt)); sh.getRange(i+1,4).setValue(salt); return {ok:true}; }
        if (action==='setRole')   { sh.getRange(i+1,6).setValue(payload.role); return {ok:true}; }
        if (action==='remove')    { sh.deleteRow(i+1); return {ok:true}; }
      }
    }
    return {error:'找不到使用者'};
  } finally { lock.releaseLock(); }
}

/** 課程頁管理：savePage（新增/修改）、setActive */
function apiAdminPages(token, action, payload) {
  requireRole_(token, 'admin');
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = ss_().getSheetByName('Pages');
    var data = sh.getDataRange().getValues();
    if (action==='savePage') {
      for (var i=1;i<data.length;i++) {
        if (data[i][0]===payload.pageId) {
          sh.getRange(i+1,2,1,6).setValues([[payload.order, payload.title, payload.subtitle, payload.contentJson, payload.questionsJson, payload.active]]);
          return {ok:true, mode:'updated'};
        }
      }
      sh.appendRow([payload.pageId, payload.order, payload.title, payload.subtitle, payload.contentJson, payload.questionsJson, payload.active]);
      return {ok:true, mode:'created'};
    }
    if (action==='setActive') {
      for (var j=1;j<data.length;j++) if (data[j][0]===payload.pageId) { sh.getRange(j+1,7).setValue(payload.active); return {ok:true}; }
      return {error:'找不到頁面'};
    }
    if (action==='raw') {
      var p = rows_('Pages').filter(function(x){ return x.pageId===payload.pageId; })[0];
      return p ? {page:p} : {error:'找不到頁面'};
    }
  } finally { lock.releaseLock(); }
}

/** 講師補充個別回饋 */
function apiAdminFeedback(token, rid, feedback) {
  requireRole_(token, 'admin');
  var sh = ss_().getSheetByName('Responses');
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++) if (data[i][0]===rid) { sh.getRange(i+1,7).setValue(feedback); return {ok:true}; }
  return {error:'找不到回答'};
}

// ====================== Web App 入口 ======================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('AI簡報力工作坊｜互動學習系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====================== 種子課程內容（依學習單） ======================
function seedPages_() {
  var sh = ss_().getSheetByName('Pages');
  if (sh.getLastRow() > 1) return; // 已有內容不覆蓋

  var visual = {
    pageId:'visual', order:1, title:'簡報視覺呈現心法', subtitle:'以認知心理學為底的四大心法',
    content: [
      {type:'lead', text:'高理解力的簡報不是「美化」，而是替觀眾的認知負荷做設計。以下四個心法，每一個都對應一個認知心理學原則。'},
      {type:'card', title:'辨識重點', body:'辨識每頁要呈現的重點內容，並排除與學習內容無關的設計。（一致性原則：無關元素會消耗工作記憶）'},
      {type:'card', title:'呈現圖像', body:'抽象內容以圖像呈現；邏輯關係以圖解呈現；細節內容或易理解概念以關鍵字呈現。（雙碼理論：圖像＋語文雙通道編碼）'},
      {type:'card', title:'重點強化', body:'訊號原則——用大小、顏色、符號強化重點。相鄰原則——圖文有關聯應擺放在相鄰位置，且應同時出現。'},
      {type:'card', title:'視覺美感', body:'字體選用非襯線字體；配色採單色的暗色與淡色組合；圖片滿版搭配關鍵字說明；圖文應對齊。'},
      {type:'links', title:'簡報素材資源', items:[
        {label:'黑色 Icon｜The Noun Project', url:'https://thenounproject.com/'},
        {label:'彩色 Icon｜Flaticon', url:'https://www.flaticon.com/'},
        {label:'圖片｜PxHere', url:'https://pxhere.com/en/'},
        {label:'圖片｜Pixabay', url:'https://pixabay.com/'},
        {label:'模板｜Slidesgo', url:'https://slidesgo.com/'},
        {label:'圖表｜PresentationGO', url:'https://www.presentationgo.com/'}
      ]}
    ],
    questions: [
      {qid:'v1', text:'請用自己的話說明：為什麼「圖文相鄰、同時出現」能幫助觀眾理解？請連結你在臨床教學的一個實際例子。',
       concepts:[{label:'相鄰原則', keys:['相鄰','鄰近','靠近','旁邊']},
                 {label:'認知負荷/工作記憶', keys:['認知負荷','工作記憶','負荷','記憶']},
                 {label:'圖文整合', keys:['圖文','整合','對應','連結']}],
       traps:[{label:'把美觀當成目的（美化≠理解）', keys:['好看','美觀','漂亮']}],
       hint:'想想觀眾的眼睛需要「來回搜尋」時，發生了什麼事？'},
      {qid:'v2', text:'一頁塞滿文字的簡報，違反了哪些心法？你會怎麼改？',
       concepts:[{label:'辨識重點', keys:['重點','取捨','刪','排除']},
                 {label:'關鍵字化', keys:['關鍵字','精簡','濃縮']},
                 {label:'訊號原則', keys:['訊號','大小','顏色','強化']}],
       traps:[{label:'只想到縮小字體而非減量', keys:['縮小','字變小']}],
       hint:'先問「這頁的唯一重點是什麼」，再決定其他內容的去留。'}
    ]
  };

  var notebook = {
    pageId:'notebook', order:2, title:'NotebookLM 應用', subtitle:'從資料整理到簡報大綱與視覺化生成',
    content: [
      {type:'lead', text:'NotebookLM（https://notebook.google.com/）讓你以「自己的資料」為邊界進行整理與生成。以下 Prompt 皆可一鍵複製，實作時依序使用。'},
      {type:'prompt', title:'① 重點摘要', body:'整理這些資料的核心訓練原則'},
      {type:'prompt', title:'② 整理內容（族群聚焦）', body:'針對 65 歲以上長者詳細完整的肌肉訓練運動規劃？'},
      {type:'prompt', title:'② 整理內容（延伸）', body:'針對 65 歲以上長者詳細完整的肌肉訓練飲食規劃？'},
      {type:'prompt', title:'③ 生成簡報大綱', body:'# 角色與任務\n請扮演一位專業且具引導力的「老年人肌力訓練教師」。請根據我所提供的資料內容，將其整理並轉化為一份「老年人肌力訓簡報大綱」。\n# 格式規範\n1. 吸晴標題：針對此教學簡報要有一個吸晴的標題。\n2. 頁數分隔：一頁內容只呈現一個重點，除非重點之間有需要相互比較再放在同一頁。每一頁簡報內容之間，請務必使用獨立一行的 `---` 符號作為分隔線。\n3. 標題格式：每頁第一行為該頁簡報標題（請勿出現「標題：」或「標題」等文字）。\n4. 內文結構：\n * 標題下方包含數個【重點關鍵字】。\n * 每個關鍵字必須單獨成段，並搭配「一句話」進行說明。\n * 「真實案例」或「引導式提問」在該關鍵字下單獨成段，加深學生的思考與印象。\n5. 最終頁面（總結）：\n * 簡報最後一頁為課程總結。\n * 請精選 1~2 句與本主題相關、簡短好記的金句或名言，作為強化核心觀念的結尾。\n6. 請**完全使用 Markdown 語法**輸出內容（包含標題 `#`、粗體 `**`、無序清單 `-` 等）。'},
      {type:'prompt', title:'④ 文件大綱 → 視覺化簡報', body:'# 角色與任務\n請扮演一位專業的「簡報視覺設計師」，請依據我提供的文件內容進行簡報視覺化設計。**依文件內容生成簡報，不要額外生成其他文字。**\n# 結構與分頁規則\n1. **分頁符號**：內容中若出現 `---` 符號，即代表分頁（切割為下一頁簡報）。\n2. **資訊階層**：請建立清晰的視覺階層（主標題、副標題、重點關鍵字）。\n3. **依序生成**：依來源內容順序依序生成簡報\n# 視覺呈現規範 (Visual & Content Rules)\n1. **圖文相鄰**：圖片/圖示若與文字內容有對應關係，請務必排版於相鄰位置。\n2. **抽象概念視覺化**：涉及「定義」或「情境描述」等抽象概念時，採用【真實照片】呈現，並於照片旁邊標示對應的【關鍵字】。\n3. **邏輯與複雜概念視覺化**：涉及「流程」、「時間軸」、「步驟」或「結構比較」等邏輯概念時，統一採用【資訊圖表 (Infographic)】方式呈現。\n4. **文字精簡**：若原始文字過多或細節過雜，請務必將其精簡濃縮為易讀的「關鍵字句」。\n# 設計系統規範 (Design System)\n1. **配色計畫 (Color Palette)**：\n * **主色 (Primary)**：藍色 `#31859C`\n * **輔助色 (Secondary)**：不同深淺的同色系藍色（用於漸層、次要區塊或背景襯底）\n * **重點強調色 (Accent)**：橘色 `#FF9900`（僅用於關鍵字、重要數據或強調框）\n * **背景色 (Background)**：純白色 `#FFFFFF`（背景不要有其他線條或圖案）\n2. **字體設定 (Typography)**：\n * 全面採用**非襯線字體 (Sans-serif)**。\n * **中文字體**：微軟正黑體 (`sans-serif`)\n * **英文字體**：`Arial`, `sans-serif`'},
      {type:'prompt', title:'⑤ 簡報內容 → 視覺化簡報（逐頁）', body:'# 角色與任務\n請扮演一位專業的「簡報視覺設計師」，請依據我提供的文件內容進行簡報視覺化設計。**依文件內容生成簡報，不要額外生成其他文字。**\n# 結構與分頁規則\n1. **依來源內容頁數**：每頁生成一頁簡報，共 OO 頁。\n2. **資訊階層**：請建立清晰的視覺階層（主標題、副標題、重點關鍵字）。\n3. **依序生成**：依來源內容順序依序生成簡報\n# 視覺呈現規範 (Visual & Content Rules)\n1. **圖文相鄰**：圖片/圖示若與文字內容有對應關係，請務必排版於相鄰位置。\n2. **抽象概念視覺化**：涉及「定義」或「情境描述」等抽象概念時，採用【真實照片】呈現，並於照片旁邊標示對應的【關鍵字】。\n3. **邏輯與複雜概念視覺化**：涉及「流程」、「時間軸」、「步驟」或「結構比較」等邏輯概念時，統一採用【資訊圖表 (Infographic)】方式呈現。\n4. **文字精簡**：若原始文字過多或細節過雜，請務必將其精簡濃縮為易讀的「關鍵字句」。\n# 設計系統規範 (Design System)\n（同上：主色 #31859C、強調 #FF9900、白底、非襯線字體）'},
      {type:'links', title:'延伸工具', items:[
        {label:'NotebookLM', url:'https://notebook.google.com/'},
        {label:'DeckEdit｜NotebookLM 簡報轉可編輯 PowerPoint', url:'https://deckedit.com/zh-Hant'}
      ]}
    ],
    questions: [
      {qid:'n1', text:'在「生成簡報大綱」的 Prompt 裡，哪些格式規範是為了降低觀眾的認知負荷？為什麼？',
       concepts:[{label:'一頁一重點', keys:['一頁','一個重點','分頁']},
                 {label:'關鍵字＋一句話', keys:['關鍵字','一句話','精簡']},
                 {label:'案例/提問促進思考', keys:['案例','提問','引導','思考']}],
       traps:[{label:'認為 AI 生成即可直接使用（未審核）', keys:['直接用','不用改','照用']}],
       hint:'對照「視覺呈現心法」那一頁的四個心法，逐條找對應。'},
      {qid:'n2', text:'NotebookLM 生成的內容，你會如何確認正確性？請寫下你自己的查核步驟。',
       concepts:[{label:'對照原始資料/引用', keys:['原始','來源','引用','出處','對照']},
                 {label:'專業判斷把關', keys:['專業','判斷','審核','查核','確認']},
                 {label:'幻覺警覺', keys:['幻覺','錯誤','虛構','編造']}],
       traps:[{label:'以流暢度判斷正確性', keys:['通順','流暢','看起來對']}],
       hint:'AI 的回答「有憑有據」和「看起來合理」是兩回事。'}
    ]
  };

  var gamma = {
    pageId:'gamma', order:3, title:'Gamma 快速產出', subtitle:'把大綱變成可編輯的視覺簡報',
    content: [
      {type:'lead', text:'Gamma 適合把已整理好的 Markdown 大綱快速轉成視覺化簡報，再匯出編修。記得：先有好大綱，才有好簡報——工具放大的是你的設計，不是取代它。'},
      {type:'links', title:'開始使用', items:[
        {label:'Gamma 邀請連結（註冊）', url:'https://gamma.app/signup?r=xtg2eu1v8uf01tc'}
      ]},
      {type:'card', title:'操作重點', body:'1. 將 NotebookLM 產出的 Markdown 大綱貼入 Gamma「貼上文字」模式。2. 選擇單欄、簡潔版型。3. 生成後逐頁檢查：重點是否單一、圖文是否相鄰、顏色是否符合你的設計系統。'},
      {type:'card', title:'匯出與編修', body:'生成結果可匯出 PPTX 繼續在 PowerPoint 微調字體（微軟正黑體）與配色（#31859C / #FF9900），維持全場作品視覺一致。'}
    ],
    questions: [
      {qid:'g1', text:'AI 生成簡報後，你「一定會手動檢查」的三件事是什麼？為什麼是這三件？',
       concepts:[{label:'內容正確性', keys:['正確','錯誤','查核','事實']},
                 {label:'一頁一重點', keys:['重點','一頁','資訊量']},
                 {label:'圖文相鄰/視覺一致', keys:['相鄰','圖文','配色','字體','一致']}],
       traps:[{label:'只檢查外觀不檢查內容', keys:['只看排版','外觀就好']}],
       hint:'回想四大心法：辨識重點、呈現圖像、重點強化、視覺美感。'}
    ]
  };

  [visual, notebook, gamma].forEach(function(p){
    sh.appendRow([p.pageId, p.order, p.title, p.subtitle, JSON.stringify(p.content), JSON.stringify(p.questions), true]);
  });
}

// ====================== 批次建立學員帳號 ======================
/**
 * 批次建立學員帳號：讀取「名單」工作表
 * 欄位：A帳號(員編) B姓名 C職類 D初始密碼(留空則用預設)
 * 已存在的帳號會自動跳過，可重複執行不會出錯
 */
function bulkCreateUsers() {
  var DEFAULT_PASS = 'cmuh0911';   // 統一初始密碼，可自行修改
  var sh = ss_().getSheetByName('名單');
  if (!sh) throw new Error('請先建立「名單」工作表，欄位：帳號/姓名/職類/初始密碼');
  var data = sh.getDataRange().getValues();
  var created = 0, skipped = 0;
  for (var i = 1; i < data.length; i++) {          // 跳過標題列
    var acc  = String(data[i][0] || '').trim();
    var name = String(data[i][1] || '').trim();
    var dept = String(data[i][2] || '').trim();
    var pass = String(data[i][3] || '').trim() || DEFAULT_PASS;
    if (!acc || !name) continue;
    if (findUserByAccount_(acc)) { skipped++; continue; }
    createUser_(acc, pass, name, 'learner', dept);
    created++;
  }
  Logger.log('建立 ' + created + ' 筆，跳過已存在 ' + skipped + ' 筆');
}

// ====================== 判燈現場控制台（講師逐題控場） ======================
/**
 * 每題四種狀態：idle（收回，學員看不到）/ open（開放作答）/ locked（鎖定，燈號凍結）/ revealed（揭曉，顯示答案分布）
 * 題目預存於 LiveBank；作答落地於 LiveAnswers。
 */
function bankSheet_() { ensureSheet_('LiveBank', ['bid','order','page','title','optionsJson','answer','explain','status']); }
function liveSheets_() {
  bankSheet_();
  ensureSheet_('LiveAnswers',['bid','uid','choice','at']);
}

function liveTally_(bid) {
  var rows = rows_('LiveAnswers').filter(function(r){ return r.bid===bid; });
  var tally = {}, total = 0;
  rows.forEach(function(r){ tally[r.choice]=(tally[r.choice]||0)+1; total++; });
  return {tally:tally, total:total};
}
function bankRows_() {
  liveSheets_();
  return rows_('LiveBank').sort(function(a,b){ return (a.order||0)-(b.order||0); });
}

/** 學員輪詢：只回傳目前 open / locked / revealed 的題目 */
function apiLiveState(token) {
  var s = requireRole_(token, null);
  var active = bankRows_().filter(function(b){ return b.status==='open'||b.status==='locked'||b.status==='revealed'; });
  if (!active.length) return {status:'idle'};
  var b = active[active.length-1];   // 以最後一題為現場題
  var mine = rows_('LiveAnswers').filter(function(r){ return r.bid===b.bid && r.uid===s.uid; })[0];
  var out = {status:b.status, bid:b.bid, page:b.page, title:b.title,
             options:JSON.parse(b.optionsJson||'[]'), myChoice: mine ? mine.choice : null};
  if (b.status === 'revealed') {
    out.answer = b.answer; out.explain = b.explain;
    var t = liveTally_(b.bid); out.tally = t.tally; out.total = t.total;
  }
  return out;
}

/** 學員作答（僅 open 狀態可送，可改選） */
function apiLiveAnswer(token, bid, choice) {
  var s = requireRole_(token, null);
  var b = bankRows_().filter(function(x){ return x.bid===bid; })[0];
  if (!b || b.status !== 'open') return {error:'本題已鎖定或尚未開放'};
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = ss_().getSheetByName('LiveAnswers');
    var data = sh.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][0]===bid && data[i][1]===s.uid) { sh.getRange(i+1,3).setValue(choice); sh.getRange(i+1,4).setValue(new Date()); return {ok:true, updated:true}; }
    }
    sh.appendRow([bid, s.uid, choice, new Date()]);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

/** 講師控制台：全題清單＋各題即時燈號分布 */
function apiAdminConsole(token) {
  requireRole_(token, 'admin');
  var enrolled = rows_('Users').filter(function(u){ return u.role==='learner'; }).length;
  var items = bankRows_().map(function(b){
    var t = liveTally_(b.bid);
    return {bid:b.bid, order:b.order, page:b.page, title:b.title,
            options:JSON.parse(b.optionsJson||'[]'), answer:b.answer, explain:b.explain,
            status:b.status||'idle', tally:t.tally, total:t.total};
  });
  var revealed = items.filter(function(i){ return i.status==='revealed'; }).length;
  var cur = items.filter(function(i){ return i.status==='open'||i.status==='locked'; }).pop() || null;
  return {items:items, enrolled:enrolled, revealed:revealed, total:items.length, current:cur};
}

/** 講師動作：setStatus（idle/open/locked/revealed）、add、update、remove */
function apiAdminBank(token, action, payload) {
  requireRole_(token, 'admin');
  liveSheets_();
  var sh = ss_().getSheetByName('LiveBank');
  if (action==='add') {
    sh.appendRow(['b'+Date.now(), payload.order||99, payload.page||'', payload.title,
      JSON.stringify(payload.options||[]), payload.answer||'', payload.explain||'', 'idle']);
    return {ok:true};
  }
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++) {
    if (data[i][0]===payload.bid) {
      if (action==='setStatus') {
        // 開放新題時，把其他仍在 open 的題目自動鎖定，避免兩題並行
        if (payload.status==='open') {
          for (var j=1;j<data.length;j++) if (j!==i && data[j][7]==='open') sh.getRange(j+1,8).setValue('locked');
        }
        sh.getRange(i+1,8).setValue(payload.status);
        return {ok:true};
      }
      if (action==='update') {
        sh.getRange(i+1,2,1,6).setValues([[payload.order, payload.page, payload.title,
          JSON.stringify(payload.options||[]), payload.answer||'', payload.explain||'']]);
        return {ok:true};
      }
      if (action==='remove') { sh.deleteRow(i+1); return {ok:true}; }
    }
  }
  return {error:'找不到題目'};
}

/** 一鍵重置全部題目為未開放（下一場工作坊前使用；不刪作答紀錄） */
function resetAllLiveStatus() {
  liveSheets_();
  var sh = ss_().getSheetByName('LiveBank');
  var n = sh.getLastRow();
  for (var i=2;i<=n;i++) sh.getRange(i,8).setValue('idle');
  Logger.log('已重置 '+(n-1)+' 題為未開放');
}

// ====================== 出口測驗（Exit Ticket） ======================
function quizSheets_() {
  ensureSheet_('Quiz',       ['qid','order','text','optionsJson','answer','active']);
  ensureSheet_('QuizAnswers',['uid','qid','choice','correct','at']);
}
function seedQuiz_() {
  quizSheets_();
  var sh = ss_().getSheetByName('Quiz');
  if (sh.getLastRow() > 1) return;
  var qs = [
    ['ex1',1,'依「相鄰原則」，圖片與對應說明文字應該如何安排？',
      ['放在相鄰位置且同時出現','圖放這頁、文字放下一頁','文字集中在最後統一說明','圖文分開以免干擾'],'放在相鄰位置且同時出現'],
    ['ex2',2,'一頁投影片的重點取捨，第一個該問的問題是？',
      ['這頁的唯一重點是什麼','還能塞進多少資訊','動畫要用哪一種','背景要選什麼顏色'],'這頁的唯一重點是什麼'],
    ['ex3',3,'「抽象概念」在簡報中建議的呈現方式是？',
      ['真實照片搭配關鍵字','大段文字詳細定義','只用口頭說明','表格逐欄列出'],'真實照片搭配關鍵字'],
    ['ex4',4,'NotebookLM 產出的內容，最重要的查核動作是？',
      ['對照原始資料來源確認','看文字通不通順','字數夠不夠多','版面漂不漂亮'],'對照原始資料來源確認'],
    ['ex5',5,'AI 生成簡報後仍必須由教師把關，主要原因是？',
      ['AI 可能產生看似合理但錯誤的內容','AI 生成速度太慢','AI 不會排版','AI 只能輸出英文'],'AI 可能產生看似合理但錯誤的內容']
  ];
  qs.forEach(function(q){ sh.appendRow([q[0], q[1], q[2], JSON.stringify(q[3]), q[4], true]); });
}

/** 學員取得測驗（不含答案）＋自己的作答結果 */
function apiQuizGet(token) {
  var s = requireRole_(token, null);
  quizSheets_();
  var qs = rows_('Quiz').filter(function(q){ return String(q.active)!=='FALSE'; })
    .sort(function(a,b){ return a.order-b.order; })
    .map(function(q){ return {qid:q.qid, text:q.text, options:JSON.parse(q.optionsJson||'[]')}; });
  var mine = rows_('QuizAnswers').filter(function(r){ return r.uid===s.uid; });
  var result = null;
  if (mine.length) {
    var key = {}; rows_('Quiz').forEach(function(q){ key[q.qid]=q.answer; });
    var details = mine.map(function(r){ return {qid:r.qid, choice:r.choice, correct:r.correct===true||r.correct==='TRUE', answer:key[r.qid]}; });
    result = {score: details.filter(function(d){return d.correct;}).length, total: details.length, details: details};
  }
  return {questions:qs, result:result};
}

/** 提交測驗（每人一次） */
function apiQuizSubmit(token, answers) {
  var s = requireRole_(token, null);
  quizSheets_();
  if (rows_('QuizAnswers').some(function(r){ return r.uid===s.uid; })) return {error:'你已提交過測驗'};
  var key = {}; rows_('Quiz').forEach(function(q){ key[q.qid]=q.answer; });
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sh = ss_().getSheetByName('QuizAnswers'), score=0, total=0;
    Object.keys(answers||{}).forEach(function(qid){
      if (!(qid in key)) return;
      var ok = String(answers[qid])===String(key[qid]);
      if (ok) score++; total++;
      sh.appendRow([s.uid, qid, String(answers[qid]), ok, new Date()]);
    });
    return {ok:true, score:score, total:total};
  } finally { lock.releaseLock(); }
}

/** 管理者：逐題答對率＋逐人成績 */
function apiAdminQuizStats(token) {
  requireRole_(token, 'admin');
  quizSheets_();
  var qs = rows_('Quiz').sort(function(a,b){ return a.order-b.order; });
  var ans = rows_('QuizAnswers');
  var users = rows_('Users');
  var byQ = qs.map(function(q){
    var a = ans.filter(function(r){ return r.qid===q.qid; });
    var ok = a.filter(function(r){ return r.correct===true||r.correct==='TRUE'; }).length;
    return {qid:q.qid, text:q.text, answered:a.length, correct:ok,
            rate: a.length ? Math.round(ok/a.length*100) : null};
  });
  var byU = {};
  ans.forEach(function(r){
    if (!byU[r.uid]) byU[r.uid]={score:0,total:0};
    byU[r.uid].total++;
    if (r.correct===true||r.correct==='TRUE') byU[r.uid].score++;
  });
  var students = Object.keys(byU).map(function(uid){
    var u = users.filter(function(x){ return x.uid===uid; })[0] || {};
    return {name:u.name||uid, dept:u.dept||'', score:byU[uid].score, total:byU[uid].total};
  }).sort(function(a,b){ return b.score-a.score; });
  return {byQuestion:byQ, students:students, submitted:students.length};
}
