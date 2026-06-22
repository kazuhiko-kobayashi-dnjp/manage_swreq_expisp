require('dotenv').config();
const express = require('express');
const fs = require('fs');
const path = require('path');
const https = require('https');
const { execFile } = require('child_process');

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, 'data', 'requirements.json');

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ========== データ読み書き ==========
function readData() {
  return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
}
function writeData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

// ========== API: 要件一覧 ==========
app.get('/api/requirements', (req, res) => {
  const data = readData();
  res.json(data);
});

// ========== API: 要件1件取得 ==========
app.get('/api/requirements/:id', (req, res) => {
  const data = readData();
  const req_item = data.req_data.find(r => r.req_id === req.params.id);
  if (!req_item) return res.status(404).json({ error: 'Not found' });
  res.json(req_item);
});

// ========== API: 要件作成 ==========
app.post('/api/requirements', async (req, res) => {
  const data = readData();
  const body = req.body;

  // req_id 自動採番
  const maxId = data.req_data.reduce((max, r) => {
    const n = parseInt(r.req_id.replace('REQ-', ''), 10);
    return n > max ? n : max;
  }, 0);
  const newId = `REQ-${String(maxId + 1).padStart(4, '0')}`;
  body.req_id = newId;
  body.created_at = new Date().toISOString();
  body.updated_at = new Date().toISOString();

  // JIRA連携（設定がある場合）
  let jiraKey = null;
  if (body.create_jira && process.env.JIRA_TOKEN && process.env.JIRA_URL) {
    try {
      jiraKey = await createJiraIssue(body);
      body.jira = body.jira || [];
      if (jiraKey && !body.jira.includes(jiraKey)) body.jira.push(jiraKey);
      body.jira_auto_created = jiraKey;
    } catch (e) {
      console.error('JIRA作成失敗:', e.message);
    }
  }

  // cat1/cat2/cat3 の順で同じグループの末尾に挿入
  const key = r => `${r.cat1||''}__${r.cat2||''}__${r.cat3||''}`;
  const bodyKey = key(body);
  let insertIdx = data.req_data.length;
  for (let i = data.req_data.length - 1; i >= 0; i--) {
    if (key(data.req_data[i]) === bodyKey) { insertIdx = i + 1; break; }
  }
  data.req_data.splice(insertIdx, 0, body);
  writeData(data);
  res.json({ ...body, jira_created: jiraKey });
});

// ========== API: 要件更新 ==========
app.put('/api/requirements/:id', async (req, res) => {
  const data = readData();
  const idx = data.req_data.findIndex(r => r.req_id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Not found' });
  const body = req.body;
  data.req_data[idx] = { ...data.req_data[idx], ...body, updated_at: new Date().toISOString() };

  // JIRA作成（更新時にチェックボックスONの場合）
  let jiraKey = null;
  if (body.create_jira && process.env.JIRA_TOKEN && process.env.JIRA_URL) {
    try {
      jiraKey = await createJiraIssue(data.req_data[idx]);
      data.req_data[idx].jira = data.req_data[idx].jira || [];
      if (jiraKey && !data.req_data[idx].jira.includes(jiraKey)) data.req_data[idx].jira.push(jiraKey);
      data.req_data[idx].jira_auto_created = jiraKey;
    } catch (e) {
      console.error('JIRA作成失敗:', e.message);
    }
  }

  writeData(data);
  res.json({ ...data.req_data[idx], jira_created: jiraKey });
});

// ========== API: 要件削除 ==========
app.delete('/api/requirements/:id', (req, res) => {
  const data = readData();
  const before = data.req_data.length;
  data.req_data = data.req_data.filter(r => r.req_id !== req.params.id);
  if (data.req_data.length === before) return res.status(404).json({ error: 'Not found' });
  writeData(data);
  res.json({ deleted: req.params.id });
});

// ========== API: イベントステータス更新（セル直接編集用）==========
app.patch('/api/requirements/:id/event', (req, res) => {
  const data = readData();
  const item = data.req_data.find(r => r.req_id === req.params.id);
  if (!item) return res.status(404).json({ error: 'Not found' });
  const { event, camera, value, status } = req.body;
  if (!item.event_plan) item.event_plan = {};
  if (!item.event_plan[event]) item.event_plan[event] = {};
  item.event_plan[event][camera] = { v: value, s: status };
  item.updated_at = new Date().toISOString();
  writeData(data);
  res.json({ ok: true, event_plan: item.event_plan });
});

// ========== API: JIRA連携テスト ==========
app.get('/api/jira/test', async (req, res) => {
  if (!process.env.JIRA_TOKEN || !process.env.JIRA_URL) {
    return res.json({ ok: false, msg: 'JIRA_TOKEN / JIRA_URL が未設定。.envファイルを確認してください。' });
  }
  try {
    const myself = await jiraGet('/rest/api/2/myself');
    res.json({ ok: true, user: myself.displayName, url: process.env.JIRA_URL });
  } catch (e) {
    res.json({ ok: false, msg: e.message });
  }
});

// ========== API: JIRAプロジェクト一覧 ==========
app.get('/api/jira/projects', async (req, res) => {
  try {
    const projects = await jiraGet('/rest/api/2/project');
    res.json(projects.map(p => ({ key: p.key, name: p.name })));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ========== API: Excel エクスポート ==========
app.get('/api/export/excel', (req, res) => {
  const script = path.join(__dirname, 'export_excel.py');
  const outFile = path.join(__dirname, 'data', '_export_tmp.xlsx');
  execFile('python3', [script, DATA_FILE, outFile], (err, stdout, stderr) => {
    if (err) {
      console.error('Excel export error:', stderr);
      return res.status(500).json({ error: stderr || err.message });
    }
    const timestamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    res.download(outFile, `BEV要件リスト_${timestamp}.xlsx`, (dlErr) => {
      if (dlErr) console.error('Download error:', dlErr);
      fs.unlink(outFile, () => {});
    });
  });
});

// ========== JIRA APIヘルパー ==========
function jiraRequest(method, path_url, body = null) {
  return new Promise((resolve, reject) => {
    const url = new URL(process.env.JIRA_URL + path_url);
    // PAT Bearer auth (Personal Access Token)
    const opts = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'Authorization': `Bearer ${process.env.JIRA_TOKEN}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
    };
    if (body) opts.headers['Content-Length'] = Buffer.byteLength(JSON.stringify(body));
    const req_h = https.request(opts, (r) => {
      let data = '';
      r.on('data', d => data += d);
      r.on('end', () => {
        try { resolve(JSON.parse(data)); } catch { resolve(data); }
      });
    });
    req_h.on('error', reject);
    if (body) req_h.write(JSON.stringify(body));
    req_h.end();
  });
}
const jiraGet = (p) => jiraRequest('GET', p);

// イベント名→納入フェーズ(customfield_14902)のIDマッピング
// BEV26EGSDプロジェクトの実際の選択肢IDから取得済み
const PHASE_OPTION_ID = {
  'VN-HILS':         68979,
  'SYS1_P1A1':       68980,
  'SYS1_P1A2':       68981,
  'SYS2_P1A1':       68982,
  'SYS2_P2A2':       68983,
  'SYS3_P1A1':       68985,
  'SYS3_P1A2':       68987,
  'SYS3_P2A2':       68988,
  'PreFOT':          68992,
  'FOT':             68990,
  'FS-CS':           68991,
  'SYS4_P1A1':       70047,
  'SYS4_P2':         73680,
  '電子CV':          70300,
  'CV':              68989,
};

// イベント名→Epic Linkマッピング（要件管理集約チケット）
const PHASE_EPIC = {
  'VN-HILS':   'BEV26EGSD-164',
  'SYS1_P1A1': 'BEV26EGSD-165',
  'SYS1_P1A2': 'BEV26EGSD-165',
  'SYS2_P1A1': 'BEV26EGSD-290',
  'SYS2_P2A2': 'BEV26EGSD-290',
  'SYS3_P1A1': 'BEV26EGSD-291',
  'SYS3_P1A2': 'BEV26EGSD-291',
  'SYS3_P2A2': 'BEV26EGSD-291',
};

async function createJiraIssue(req_item) {
  const project = process.env.JIRA_PROJECT || 'BEV26EGSD';

  // summary: [機能名]要件名 形式
  const funcName = req_item.cat2 || req_item.cat1 || '';
  const summary = `[${funcName}]${req_item.desc}`;

  // description: テンプレ構造に準拠
  const cat1 = req_item.cat1 || '';
  const cat2 = req_item.cat2 || '';
  const cat3short = (req_item.cat3 || '').replace(/^MVH-VDCTRL_/, '');
  const description = [
    '■要件概要',
    req_item.summary || req_item.desc || '',
    '',
    '■仕様書情報',
    `仕様書: ${req_item.cat3 || ''}`,
    `版数: ${req_item.ver || ''}`,
    `発行日: ${req_item.issue_date || ''}`,
    `対象カメラ: ${(req_item.cameras || []).join(', ')}`,
    `SW挙動変更: ${req_item.sw || ''}`,
    '',
    '■ADC要件リスト',
    'ADC要件リストに要件項目を追加します。',
    '||機能大分類||中分類||小分類||',
    `|${cat1}|${cat2}|${cat3short}|`,
    '',
    `管理ID: ${req_item.req_id || ''}`,
  ].join('\n');

  // 最初にイベント計画があるフェーズを納入フェーズとして設定
  let phaseOptionId = null;
  let epicLink = null;
  const ep = req_item.event_plan || {};
  for (const evName of Object.keys(PHASE_OPTION_ID)) {
    if (ep[evName] && Object.values(ep[evName]).some(ci => ci && ci.v && ci.v !== '-')) {
      phaseOptionId = PHASE_OPTION_ID[evName];
      epicLink = PHASE_EPIC[evName] || null;
      break;
    }
  }
  // フォームから直接指定された場合は上書き
  if (req_item.jira_phase) {
    phaseOptionId = PHASE_OPTION_ID[req_item.jira_phase] || phaseOptionId;
    epicLink = PHASE_EPIC[req_item.jira_phase] || epicLink;
  }

  // ラベル: sw影響 + 担当者所属（req_item.subteamやpersonsから）
  const labels = [];
  if (req_item.sw === 'Y') labels.push('ソフト影響あり');
  else if (req_item.sw === 'N') labels.push('ソフト影響なし');

  const fields = {
    project: { key: project },
    summary,
    description,
    issuetype: { id: '10205' },  // ストーリー
    labels,
  };

  // 納入フェーズ（選択肢）
  if (phaseOptionId) {
    fields['customfield_14902'] = { id: String(phaseOptionId) };
  }
  // Epic Link
  if (epicLink) {
    fields['customfield_10100'] = epicLink;
  }
  // 開始日（明示指定があればそちら、なければ発行日）
  const startDate = req_item.jira_start || req_item.issue_date;
  if (startDate) {
    fields['customfield_10200'] = startDate;
  }
  // 期限
  if (req_item.jira_due) {
    fields['duedate'] = req_item.jira_due;
  }
  // 担当者（仕様担当）
  if (req_item.jira_assignee) {
    fields['assignee'] = { name: req_item.jira_assignee };
  }

  const result = await jiraRequest('POST', '/rest/api/2/issue', { fields });
  if (result.errors) throw new Error(JSON.stringify(result.errors));
  return result.key;
}

// ========== API: イベント設定取得・更新 ==========
app.get('/api/events', (req, res) => {
  const data = readData();
  res.json(data.events_config || []);
});

app.put('/api/events', (req, res) => {
  const data = readData();
  data.events_config = req.body;
  data.events = req.body.map(e => e.name);
  writeData(data);
  res.json(data.events_config);
});

// ========== 404フォールバック（SPA） ==========
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`\n🚀 BEV要件管理サーバー起動`);
  console.log(`   http://localhost:${PORT}`);
  console.log(`   データ: ${DATA_FILE}`);
  const jiraOk = !!(process.env.JIRA_TOKEN && process.env.JIRA_URL);
  console.log(`   JIRA連携: ${jiraOk ? '✅ 有効' : '❌ 無効 (.envを設定してください)'}`);
  console.log('');
});
