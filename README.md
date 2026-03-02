<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>UKETUGI 議事録生成ツール</title>

  <!-- ===== 設定ファイル（APIキー等） ===== -->
  <!-- config.example.js をコピーして config.js を作成してください -->
  <script src="config.js"></script>

  <!-- ===== 外部ライブラリ（CDN） ===== -->
  <!-- docx.js: ブラウザ上でWord文書(.docx)を生成するライブラリ -->
  <script src="https://unpkg.com/docx@8.5.0/build/index.umd.js"></script>
  <!-- FileSaver.js: 生成したファイルをダウンロードさせるライブラリ -->
  <script src="https://unpkg.com/file-saver@2.0.5/dist/FileSaver.min.js"></script>
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700&display=swap" rel="stylesheet">

  <style>
    /* ===== リセット & ベース ===== */
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --navy: #2E4057;
      --navy-light: #3D5470;
      --accent: #3B82F6;
      --accent-hover: #2563EB;
      --bg: #F0F2F5;
      --card: #FFFFFF;
      --text: #1A1A2E;
      --text-sub: #6B7280;
      --border: #E5E7EB;
      --success: #10B981;
      --error: #EF4444;
      --radius: 12px;
    }

    body {
      font-family: 'Noto Sans JP', sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.7;
      min-height: 100vh;
    }

    /* ===== ヘッダー ===== */
    .header {
      background: linear-gradient(135deg, var(--navy) 0%, var(--navy-light) 100%);
      padding: 28px 32px;
      color: white;
      position: relative;
      overflow: hidden;
    }
    .header::after {
      content: '';
      position: absolute;
      top: -50%;
      right: -10%;
      width: 400px;
      height: 400px;
      background: radial-gradient(circle, rgba(255,255,255,0.06) 0%, transparent 70%);
      border-radius: 50%;
    }
    .header h1 {
      font-family: 'Outfit', sans-serif;
      font-size: 26px;
      font-weight: 700;
      letter-spacing: 1px;
    }
    .header p {
      font-size: 13px;
      opacity: 0.75;
      margin-top: 4px;
    }

    /* ===== メインコンテナ ===== */
    .container {
      max-width: 760px;
      margin: 0 auto;
      padding: 32px 20px 60px;
    }

    /* ===== カード ===== */
    .card {
      background: var(--card);
      border-radius: var(--radius);
      padding: 28px;
      margin-bottom: 20px;
      border: 1px solid var(--border);
      box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .card-title {
      font-size: 15px;
      font-weight: 600;
      color: var(--navy);
      margin-bottom: 20px;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .card-title .icon {
      width: 28px;
      height: 28px;
      background: var(--navy);
      color: white;
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      font-family: 'Outfit', sans-serif;
      font-weight: 600;
    }

    /* ===== フォーム要素 ===== */
    .form-group {
      margin-bottom: 18px;
    }
    .form-group:last-child {
      margin-bottom: 0;
    }
    label {
      display: block;
      font-size: 13px;
      font-weight: 500;
      color: var(--text-sub);
      margin-bottom: 6px;
    }
    input[type="text"],
    input[type="password"],
    input[type="datetime-local"],
    textarea {
      width: 100%;
      padding: 10px 14px;
      border: 1.5px solid var(--border);
      border-radius: 8px;
      font-family: 'Noto Sans JP', sans-serif;
      font-size: 14px;
      color: var(--text);
      transition: border-color 0.2s;
      background: #FAFBFC;
    }
    input:focus, textarea:focus {
      outline: none;
      border-color: var(--accent);
      background: white;
    }
    textarea {
      resize: vertical;
      min-height: 80px;
    }
    .hint {
      font-size: 11px;
      color: var(--text-sub);
      margin-top: 4px;
    }

    /* ===== ファイルアップロード ===== */
    .file-upload {
      border: 2px dashed var(--border);
      border-radius: var(--radius);
      padding: 36px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s;
      background: #FAFBFC;
      position: relative;
    }
    .file-upload:hover {
      border-color: var(--accent);
      background: #F0F7FF;
    }
    .file-upload.has-file {
      border-color: var(--success);
      background: #F0FDF4;
    }
    .file-upload input[type="file"] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
    }
    .file-upload .upload-icon {
      font-size: 32px;
      margin-bottom: 8px;
    }
    .file-upload .upload-text {
      font-size: 14px;
      color: var(--text-sub);
    }
    .file-upload .file-name {
      font-size: 14px;
      color: var(--success);
      font-weight: 500;
      margin-top: 4px;
    }

    /* ===== ボタン ===== */
    .btn-primary {
      width: 100%;
      padding: 14px 24px;
      background: linear-gradient(135deg, var(--accent) 0%, var(--accent-hover) 100%);
      color: white;
      border: none;
      border-radius: 10px;
      font-family: 'Noto Sans JP', sans-serif;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      letter-spacing: 0.5px;
    }
    .btn-primary:hover:not(:disabled) {
      transform: translateY(-1px);
      box-shadow: 0 4px 14px rgba(59, 130, 246, 0.35);
    }
    .btn-primary:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    /* ===== プログレス ===== */
    .progress-area {
      display: none;
      margin-top: 24px;
    }
    .progress-area.active {
      display: block;
    }
    .progress-bar-track {
      height: 6px;
      background: var(--border);
      border-radius: 3px;
      overflow: hidden;
      margin-bottom: 12px;
    }
    .progress-bar-fill {
      height: 100%;
      background: linear-gradient(90deg, var(--accent), #818CF8);
      border-radius: 3px;
      transition: width 0.5s ease;
      width: 0%;
    }
    .progress-status {
      font-size: 13px;
      color: var(--text-sub);
      text-align: center;
    }

    /* ===== 結果エリア ===== */
    .result-area {
      display: none;
      text-align: center;
      padding: 24px;
    }
    .result-area.active {
      display: block;
    }
    .result-area .success-icon {
      font-size: 48px;
      margin-bottom: 12px;
    }
    .result-area p {
      color: var(--text-sub);
      font-size: 14px;
    }

    /* ===== エラー ===== */
    .error-msg {
      display: none;
      background: #FEF2F2;
      border: 1px solid #FECACA;
      border-radius: 8px;
      padding: 12px 16px;
      color: var(--error);
      font-size: 13px;
      margin-top: 12px;
    }
    .error-msg.active {
      display: block;
    }

    /* ===== 設定バナー ===== */
    .config-banner {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 12px 16px;
      border-radius: var(--radius);
      font-size: 13px;
      margin-bottom: 20px;
      background: #F0FDF4;
      border: 1px solid #BBF7D0;
      color: #166534;
    }
    .config-banner.error {
      background: #FEF2F2;
      border: 1px solid #FECACA;
      color: #991B1B;
    }
    .config-icon {
      font-size: 16px;
      flex-shrink: 0;
    }

    /* ===== フッター ===== */
    .footer {
      text-align: center;
      padding: 20px;
      font-size: 11px;
      color: var(--text-sub);
    }
  </style>
</head>
<body>

  <!-- ===== ヘッダー ===== -->
  <div class="header">
    <h1>UKETUGI Minutes</h1>
    <p>文字起こしデータから議事録を自動生成</p>
  </div>

  <div class="container">

    <!-- ===== 設定状態バナー ===== -->
    <div class="config-banner" id="configBanner">
      <span class="config-icon">⚙️</span>
      <span id="configStatus">設定を確認中...</span>
    </div>

    <!-- ===== 会議情報 ===== -->
    <div class="card">
      <div class="card-title">
        <div class="icon">1</div>
        会議情報
      </div>
      <div class="form-group">
        <label>会議名</label>
        <input type="text" id="meetingName" placeholder="例: UKETUGI 定例会議" value="UKETUGI 定例会議">
      </div>
      <div class="form-group">
        <label>開催日時</label>
        <input type="datetime-local" id="meetingDate">
      </div>
      <div class="form-group">
        <label>参加者（カンマ区切り）</label>
        <input type="text" id="participants" placeholder="例: 黒越、小林、佐藤、伊藤、長井">
      </div>
      <div class="form-group">
        <label>場所・形式</label>
        <input type="text" id="meetingLocation" placeholder="例: オンライン会議" value="オンライン会議">
      </div>
    </div>

    <!-- ===== 文字起こしアップロード ===== -->
    <div class="card">
      <div class="card-title">
        <div class="icon">2</div>
        文字起こしデータ
      </div>
      <div class="file-upload" id="fileUploadArea">
        <div class="upload-icon">📄</div>
        <div class="upload-text">クリックまたはドラッグ＆ドロップで<br>テキストファイル(.txt)をアップロード</div>
        <div class="file-name" id="fileName"></div>
        <input type="file" id="fileInput" accept=".txt,.text">
      </div>
      <p class="hint" style="margin-top:8px">または下のテキストエリアに直接ペースト</p>
      <div class="form-group" style="margin-top:12px">
        <textarea id="transcriptText" rows="6" placeholder="ここに文字起こしデータを貼り付け..."></textarea>
      </div>
    </div>

    <!-- ===== 生成ボタン ===== -->
    <button class="btn-primary" id="generateBtn" onclick="handleGenerate()">
      議事録を生成する
    </button>

    <!-- ===== エラー表示 ===== -->
    <div class="error-msg" id="errorMsg"></div>

    <!-- ===== 進捗表示 ===== -->
    <div class="progress-area" id="progressArea">
      <div class="progress-bar-track">
        <div class="progress-bar-fill" id="progressFill"></div>
      </div>
      <div class="progress-status" id="progressStatus">準備中...</div>
    </div>

    <!-- ===== 結果表示 ===== -->
    <div class="result-area card" id="resultArea">
      <div class="success-icon">✅</div>
      <p><strong>議事録の生成が完了しました！</strong></p>
      <p style="margin-top:8px">ファイルが自動でダウンロードされます。<br>されない場合は下のボタンをクリックしてください。</p>
      <button class="btn-primary" style="margin-top:16px; max-width:300px; margin-left:auto; margin-right:auto;" onclick="downloadAgain()">
        再ダウンロード
      </button>
    </div>

  </div>

  <div class="footer">
    UKETUGI Minutes Tool — Powered by Claude API + docx.js
  </div>

<script>
// =============================================================================
// グローバル変数
// =============================================================================
let transcriptContent = '';   // 文字起こしテキスト
let generatedBuffer = null;   // 生成済みdocxバッファ
let generatedFileName = '';   // 生成済みファイル名

// =============================================================================
// UI操作
// =============================================================================

// 設定ファイル状態チェック（ページ読み込み時）
function checkConfig() {
  const banner = document.getElementById('configBanner');
  const status = document.getElementById('configStatus');

  if (typeof APP_CONFIG === 'undefined') {
    banner.classList.add('error');
    status.textContent = '⚠ config.js が見つかりません。config.example.js をコピーして config.js を作成してください。';
    document.getElementById('generateBtn').disabled = true;
    return;
  }
  if (!APP_CONFIG.PROXY_URL || APP_CONFIG.PROXY_URL.includes('your-')) {
    banner.classList.add('error');
    status.textContent = '⚠ config.js のプロキシURLが未設定です。';
    document.getElementById('generateBtn').disabled = true;
    return;
  }
  if (!APP_CONFIG.ACCESS_TOKEN || APP_CONFIG.ACCESS_TOKEN === '') {
    banner.classList.add('error');
    status.textContent = '⚠ config.js の合言葉（ACCESS_TOKEN）が未設定です。';
    document.getElementById('generateBtn').disabled = true;
    return;
  }

  banner.classList.remove('error');
  status.textContent = '✅ 設定OK — API接続準備完了';
}

// ファイルアップロード処理
document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    transcriptContent = ev.target.result;
    document.getElementById('fileName').textContent = '📎 ' + file.name;
    document.getElementById('fileUploadArea').classList.add('has-file');
    document.getElementById('transcriptText').value = transcriptContent;
  };
  reader.readAsText(file, 'UTF-8');
});

// プログレス更新
function setProgress(pct, msg) {
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressStatus').textContent = msg;
}

// エラー表示
function showError(msg) {
  const el = document.getElementById('errorMsg');
  el.textContent = msg;
  el.classList.add('active');
}
function hideError() {
  document.getElementById('errorMsg').classList.remove('active');
}

// 再ダウンロード
function downloadAgain() {
  if (generatedBuffer) {
    saveAs(generatedBuffer, generatedFileName);
  }
}

// =============================================================================
// メイン処理: 議事録生成
// =============================================================================
async function handleGenerate() {
  hideError();

  // --- config.js チェック ---
  if (typeof APP_CONFIG === 'undefined' || !APP_CONFIG.ACCESS_TOKEN) {
    showError('config.js が未設定です。config.example.js をコピーして合言葉を設定してください。');
    return;
  }

  const proxyUrl = APP_CONFIG.PROXY_URL || '';
  const accessToken = APP_CONFIG.ACCESS_TOKEN || '';
  const model = APP_CONFIG.MODEL || 'claude-sonnet-4-20250514';

  // --- バリデーション ---
  const meetingName = document.getElementById('meetingName').value.trim();
  const meetingDate = document.getElementById('meetingDate').value;
  const participants = document.getElementById('participants').value.trim();
  const meetingLocation = document.getElementById('meetingLocation').value.trim();
  const transcript = document.getElementById('transcriptText').value.trim() || transcriptContent;

  if (!meetingName) { showError('会議名を入力してください'); return; }
  if (!meetingDate) { showError('開催日時を入力してください'); return; }
  if (!participants) { showError('参加者を入力してください'); return; }
  if (!transcript) { showError('文字起こしデータをアップロードまたはペーストしてください'); return; }

  // --- UI更新 ---
  const btn = document.getElementById('generateBtn');
  btn.disabled = true;
  btn.textContent = '生成中...';
  document.getElementById('progressArea').classList.add('active');
  document.getElementById('resultArea').classList.remove('active');

  try {
    // ステップ1: Claude APIで構造化
    setProgress(10, 'AIが文字起こしを分析中...');
    const structured = await callClaudeAPI(proxyUrl, accessToken, model, transcript, meetingName, participants);

    // ステップ2: マインドマップSVG生成
    setProgress(50, 'マインドマップを生成中...');
    const mindmapPng = await generateMindmapPng(structured.keywords, meetingName, meetingDate);

    // ステップ3: docx生成
    setProgress(70, 'Word文書を生成中...');
    const dateObj = new Date(meetingDate);
    const dateStr = formatDateJa(dateObj);
    const buffer = await generateDocx(structured, meetingName, dateStr, participants, meetingLocation, mindmapPng);

    // ステップ4: ダウンロード
    setProgress(95, 'ダウンロード準備中...');
    const fileDate = meetingDate.replace(/[-T:]/g, '').slice(0, 8);
    generatedFileName = fileDate + '_' + meetingName.replace(/\s+/g, '_') + '_議事録.docx';
    generatedBuffer = buffer;

    saveAs(buffer, generatedFileName);

    setProgress(100, '完了！');
    document.getElementById('resultArea').classList.add('active');

  } catch (err) {
    console.error(err);
    showError('エラーが発生しました: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = '議事録を生成する';
  }
}

// =============================================================================
// Claude API呼び出し
// =============================================================================
async function callClaudeAPI(proxyUrl, accessToken, model, transcript, meetingName, participants) {
  const endpoint = proxyUrl;

  const systemPrompt = `あなたは議事録作成の専門家です。会議の文字起こしデータを受け取り、構造化された議事録データをJSON形式で返してください。

以下のJSON構造で返してください。JSON以外のテキスト（説明文やマークダウンのバッククォート）は一切含めないでください:

{
  "executive_summary": "会議全体の要点を3〜5行で要約した文章",
  "decisions": ["決定事項1", "決定事項2", ...],
  "topics": [
    {
      "title": "議題名",
      "sections": [
        { "label": "背景・課題", "points": ["ポイント1", "ポイント2"] },
        { "label": "議論内容", "points": ["ポイント1", "ポイント2"] },
        { "label": "結論", "points": ["ポイント1", "ポイント2"] }
      ]
    }
  ],
  "keywords": {
    "カテゴリ名1": { "color": "DBEAFE", "keywords": ["キーワード1", "キーワード2"] },
    "カテゴリ名2": { "color": "D1FAE5", "keywords": ["キーワード1", "キーワード2"] }
  },
  "action_items": [
    { "task": "タスク内容", "assignee": "担当者", "deadline": "期限", "status": "ステータス" }
  ],
  "next_meeting": {
    "date_note": "次回日程に関するメモ",
    "agenda_candidates": ["議題候補1", "議題候補2"]
  }
}

キーワードのカテゴリは4〜6個作成し、各カテゴリに適切な背景色(HEX、明るい色)を設定してください。色の候補: DBEAFE(青), D1FAE5(緑), FEE2E2(赤), CFFAFE(水色), FEF3C7(黄), EDE9FE(紫), FDE68A(オレンジ), E0E7FF(インディゴ)

議題(topics)はセクション分けを柔軟に行い、内容に応じて「背景・課題」「議論内容」「結論」「現状」「戦略」「検討ツール」など適切なラベルを使ってください。

action_itemsのstatusは「着手予定」「未着手」「進行中」「継続」「テスト中」「翌日実施」等から適切なものを選んでください。`;

  const userPrompt = `以下は「${meetingName}」（参加者: ${participants}）の文字起こしデータです。この内容から構造化された議事録JSONを生成してください。

--- 文字起こしデータ ---
${transcript}
--- ここまで ---

JSONのみを出力してください。`;

  const body = {
    model: model,
    max_tokens: 4000,
    messages: [
      { role: 'user', content: userPrompt }
    ],
    system: systemPrompt
  };

  const headers = {
    'Content-Type': 'application/json',
    'x-access-token': accessToken,
    'anthropic-version': '2023-06-01'
  };

  const res = await fetch(endpoint, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(body)
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`API Error (${res.status}): ${errText}`);
  }

  const data = await res.json();
  const text = data.content
    .filter(c => c.type === 'text')
    .map(c => c.text)
    .join('');

  // JSONパース（バッククォート除去）
  const cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    console.error('JSON parse error. Raw text:', text);
    throw new Error('AIの応答をJSONとして解析できませんでした。再試行してください。');
  }
}

// =============================================================================
// マインドマップ生成（SVG → PNG）
// =============================================================================
async function generateMindmapPng(keywords, meetingName, meetingDate) {
  const categories = Object.entries(keywords);
  const catCount = categories.length;

  // カテゴリごとの色（ノード本体の色）
  const branchColors = ['#3B82F6', '#10B981', '#F59E0B', '#8B5CF6', '#EF4444', '#06B6D4', '#F97316', '#6366F1'];

  const W = 1200, H = 820;
  const cx = W / 2, cy = H / 2;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">`;
  svg += `<defs><style>text { font-family: Arial, Helvetica, sans-serif; }</style></defs>`;
  svg += `<rect width="${W}" height="${H}" fill="#FAFBFC" rx="8"/>`;

  // カテゴリ配置計算（中央から放射状）
  const positions = [];
  for (let i = 0; i < catCount; i++) {
    const angle = (2 * Math.PI * i / catCount) - Math.PI / 2;
    const rx = 280, ry = 200;
    const bx = cx + rx * Math.cos(angle);
    const by = cy + ry * Math.sin(angle);
    positions.push({ bx, by, angle });
  }

  // 接続線（中央→カテゴリ）
  categories.forEach(([, ], i) => {
    const { bx, by } = positions[i];
    const color = branchColors[i % branchColors.length];
    const mcx = (cx + bx) / 2;
    const mcy = cy + (by - cy) * 0.3;
    svg += `<path d="M ${cx},${cy} Q ${mcx},${mcy} ${bx},${by}" stroke="${color}" stroke-width="2.5" fill="none" opacity="0.45"/>`;
  });

  // カテゴリノード＋サブキーワード
  categories.forEach(([catName, catData], i) => {
    const { bx, by, angle } = positions[i];
    const color = branchColors[i % branchColors.length];
    const bgColor = '#' + (catData.color || 'E5E7EB');
    const kws = catData.keywords || [];

    // カテゴリノード
    const labelW = Math.max(catName.length * 16, 120);
    svg += `<rect x="${bx - labelW/2}" y="${by - 22}" width="${labelW}" height="44" rx="22" fill="${color}"/>`;
    svg += `<text x="${bx}" y="${by + 5}" text-anchor="middle" fill="white" font-size="13" font-weight="bold">${escXml(catName)}</text>`;

    // サブキーワード（カテゴリの外側に配置）
    const kwCount = Math.min(kws.length, 5); // 最大5個表示
    kws.slice(0, kwCount).forEach((kw, j) => {
      const subAngle = angle + (j - (kwCount - 1) / 2) * 0.3;
      const subR = 140 + (j % 2) * 20;
      const sx = bx + subR * Math.cos(subAngle);
      const sy = by + subR * Math.sin(subAngle);

      // 接続線
      svg += `<line x1="${bx}" y1="${by}" x2="${sx}" y2="${sy}" stroke="${color}" stroke-width="1.2" opacity="0.3"/>`;

      // サブノード
      const kwW = Math.max(kw.length * 13, 70);
      svg += `<rect x="${sx - kwW/2}" y="${sy - 14}" width="${kwW}" height="28" rx="14" fill="${bgColor}" stroke="${color}" stroke-width="1"/>`;
      svg += `<text x="${sx}" y="${sy + 4}" text-anchor="middle" fill="${color}" font-size="10">${escXml(kw)}</text>`;
    });
  });

  // 中央ノード
  svg += `<rect x="${cx - 90}" y="${cy - 32}" width="180" height="64" rx="32" fill="#2E4057"/>`;
  svg += `<text x="${cx}" y="${cy - 4}" text-anchor="middle" fill="white" font-size="16" font-weight="bold">${escXml(meetingName.length > 12 ? meetingName.slice(0,12)+'...' : meetingName)}</text>`;

  const dateLabel = meetingDate ? new Date(meetingDate).toLocaleDateString('ja-JP') : '';
  if (dateLabel) {
    svg += `<text x="${cx}" y="${cy + 16}" text-anchor="middle" fill="#CBD5E1" font-size="10">${escXml(dateLabel)}</text>`;
  }

  svg += `</svg>`;

  // SVG → PNG（Canvas経由）
  return new Promise((resolve, reject) => {
    const img = new Image();
    const svgBlob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(svgBlob);

    img.onload = function() {
      const canvas = document.createElement('canvas');
      canvas.width = W * 2;  // 高解像度
      canvas.height = H * 2;
      const ctx = canvas.getContext('2d');
      ctx.scale(2, 2);
      ctx.drawImage(img, 0, 0, W, H);
      URL.revokeObjectURL(url);

      canvas.toBlob(function(blob) {
        const reader = new FileReader();
        reader.onload = () => resolve(new Uint8Array(reader.result));
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob);
      }, 'image/png');
    };
    img.onerror = reject;
    img.src = url;
  });
}

function escXml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// =============================================================================
// docx生成
// =============================================================================
async function generateDocx(data, meetingName, dateStr, participants, location, mindmapPng) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    LevelFormat, PageBreak
  } = docx;

  const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cm = { top: 60, bottom: 60, left: 100, right: 100 }; // cell margins

  // --- ヘルパー関数 ---
  function hCell(text, w) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA },
      shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: cm,
      children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20, bold: true, color: "FFFFFF" })] })]
    });
  }
  function dCell(text, w, opts = {}) {
    return new TableCell({
      borders, width: { size: w, type: WidthType.DXA },
      shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
      margins: cm,
      children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20, bold: !!opts.bold })] })]
    });
  }
  function bullet(text) {
    return new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(text)] });
  }
  function sectionLabel(text) {
    return new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: '【' + text + '】', bold: true })] });
  }

  // --- children配列を組み立て ---
  const children = [];

  // タイトル
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
    children: [new TextRun({ text: "議事録", font: "Arial", size: 40, bold: true, color: "2E4057" })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({ text: meetingName, font: "Arial", size: 28, color: "555555" })] }));

  // 基本情報
  children.push(new Table({
    width: { size: 9506, type: WidthType.DXA }, columnWidths: [2400, 7106],
    rows: [
      new TableRow({ children: [dCell("日時", 2400, { fill: "E8EDF2", bold: true }), dCell(dateStr, 7106)] }),
      new TableRow({ children: [dCell("場所", 2400, { fill: "E8EDF2", bold: true }), dCell(location, 7106)] }),
      new TableRow({ children: [dCell("参加者", 2400, { fill: "E8EDF2", bold: true }), dCell(participants, 7106)] }),
      new TableRow({ children: [dCell("記録者", 2400, { fill: "E8EDF2", bold: true }), dCell("AI自動生成", 7106)] }),
    ]
  }));
  children.push(new Paragraph({ spacing: { before: 300 } }));

  // エグゼクティブサマリー
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("エグゼクティブサマリー")] }));
  children.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun(data.executive_summary || '')] }));

  // キーワード一覧
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("キーワード一覧")] }));
  const kwRows = [new TableRow({ children: [hCell("カテゴリ", 2200), hCell("キーワード", 7306)] })];
  Object.entries(data.keywords || {}).forEach(([cat, catData]) => {
    const kws = (catData.keywords || []).join(' ／ ');
    kwRows.push(new TableRow({ children: [
      dCell(cat, 2200, { fill: catData.color || "E5E7EB", bold: true }),
      dCell(kws, 7306)
    ]}));
  });
  children.push(new Table({ width: { size: 9506, type: WidthType.DXA }, columnWidths: [2200, 7306], rows: kwRows }));

  // マインドマップ
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("キーワード マインドマップ")] }));
  children.push(new Paragraph({ spacing: { after: 60 }, children: [
    new TextRun({ text: "会議で議論された主要トピックとキーワードの関連を以下に可視化。", font: "Arial", size: 20, color: "555555" })
  ]}));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new ImageRun({ data: mindmapPng, transformation: { width: 700, height: 478 }, type: "png" })
  ]}));
  children.push(new Paragraph({ spacing: { before: 200 } }));

  // 決定事項
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("決定事項")] }));
  (data.decisions || []).forEach(d => children.push(bullet(d)));

  // 議論の経緯
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("議論の経緯")] }));
  (data.topics || []).forEach((topic, i) => {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2,
      children: [new TextRun(`議題${i + 1}: ${topic.title}`)] }));
    (topic.sections || []).forEach(sec => {
      children.push(sectionLabel(sec.label));
      (sec.points || []).forEach(p => children.push(bullet(p)));
    });
  });

  // アクションアイテム
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("アクションアイテム")] }));
  const aiRows = [new TableRow({ children: [
    hCell("No", 600), hCell("タスク", 4400), hCell("担当", 1400), hCell("期限", 1200), hCell("ステータス", 1360)
  ]})];
  (data.action_items || []).forEach((item, i) => {
    aiRows.push(new TableRow({ children: [
      dCell(String(i + 1), 600), dCell(item.task, 4400), dCell(item.assignee, 1400),
      dCell(item.deadline, 1200), dCell(item.status, 1360)
    ]}));
  });
  children.push(new Table({ width: { size: 9506, type: WidthType.DXA }, columnWidths: [600, 4400, 1400, 1200, 1360], rows: aiRows }));
  children.push(new Paragraph({ spacing: { before: 300 } }));

  // 次回予定
  if (data.next_meeting) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("次回予定")] }));
    if (data.next_meeting.date_note) {
      children.push(bullet('日時: ' + data.next_meeting.date_note));
    }
    if (data.next_meeting.agenda_candidates) {
      children.push(bullet('議題候補: ' + data.next_meeting.agenda_candidates.join('、')));
    }
  }

  // 以上
  children.push(new Paragraph({ spacing: { before: 400 } }));
  children.push(new Paragraph({ alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text: "以上", font: "Arial", size: 22, color: "555555" })] }));

  // --- Document生成 ---
  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 32, bold: true, font: "Arial", color: "2E4057" },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, font: "Arial", color: "2E4057" },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      ]
    },
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
        }
      },
      children
    }]
  });

  return await Packer.toBlob(doc);
}

// =============================================================================
// ユーティリティ
// =============================================================================
function formatDateJa(d) {
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const dow = days[d.getDay()];
  const h = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${y}年${m}月${day}日（${dow}）${h}:${min}〜`;
}

// 初期化
(function init() {
  const now = new Date();
  const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  document.getElementById('meetingDate').value = local.toISOString().slice(0, 16);
  checkConfig();
})();
</script>

</body>
</html>
