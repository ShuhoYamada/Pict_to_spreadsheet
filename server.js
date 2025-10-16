const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// デバッグ: 環境変数の確認
console.log('🔍 環境変数チェック:');
console.log('GOOGLE_CLIENT_ID:', process.env.GOOGLE_CLIENT_ID ? 'あり' : 'なし');
console.log('GOOGLE_CLIENT_SECRET:', process.env.GOOGLE_CLIENT_SECRET ? 'あり' : 'なし');
console.log('PORT:', PORT);

// ミドルウェア
app.use(cors());
app.use(express.json());
app.use(express.static('public')); // フロントエンドファイルを提供

// multerの設定（メモリーストレージ）
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Google APIs設定
console.log('🔧 OAuth2Client初期化中...');
if (!process.env.GOOGLE_CLIENT_ID || !process.env.GOOGLE_CLIENT_SECRET) {
  console.error('❌ 必要な環境変数が設定されていません!');
  console.error('GOOGLE_CLIENT_ID:', process.env.GOOGLE_CLIENT_ID);
  console.error('GOOGLE_CLIENT_SECRET:', process.env.GOOGLE_CLIENT_SECRET);
  process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  'http://localhost:3000/auth/callback'
);

// OAuth2Clientの設定を確認
console.log('✅ OAuth2Client初期化完了');
console.log('🔍 OAuth2Client設定確認:');
console.log('Client ID:', oauth2Client._clientId);
console.log('Client Secret:', oauth2Client._clientSecret ? 'あり' : 'なし');
console.log('Redirect URI:', oauth2Client.redirectUri);

// Google Drive API
const drive = google.drive({ version: 'v3', auth: oauth2Client });
const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

// ヘルスチェックエンドポイント
app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'セキュアサーバーが正常に動作中' });
});

// 認証URL生成
app.get('/auth/url', (req, res) => {
  try {
    console.log('🔐 認証URL生成リクエストを受信');
    
    const scopes = [
      'https://www.googleapis.com/auth/drive.readonly',
      'https://www.googleapis.com/auth/spreadsheets'
    ];
    
    const authUrl = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: scopes.join(' '),
      response_type: 'code',
      client_id: process.env.GOOGLE_CLIENT_ID,
      redirect_uri: 'http://localhost:3000/auth/callback'
    });
    
    console.log('🔍 認証URL詳細確認:');
    console.log('Client ID:', process.env.GOOGLE_CLIENT_ID);
    console.log('Redirect URI: http://localhost:3000/auth/callback');
    
    console.log('✅ 認証URL生成成功:', authUrl);
    res.json({ authUrl });
  } catch (error) {
    console.error('❌ 認証URL生成エラー:', error);
    res.status(500).json({ error: '認証URL生成に失敗しました', details: error.message });
  }
});

// 認証コールバック
app.get('/auth/callback', async (req, res) => {
  const { code, error } = req.query;
  
  console.log('🔄 認証コールバック受信:', { code: code ? 'あり' : 'なし', error });
  
  if (error) {
    console.error('❌ OAuth認証エラー:', error);
    return res.redirect(`/?auth=error&message=${encodeURIComponent(error)}`);
  }
  
  if (!code) {
    console.error('❌ 認証コードが見つかりません');
    return res.redirect('/?auth=error&message=no_code');
  }
  
  try {
    console.log('🔑 アクセストークン取得中...');
    
    // 正しいメソッドを使用してトークンを取得
    const tokenResponse = await oauth2Client.getToken(code);
    console.log('🔍 トークン取得レスポンス:', {
      tokens: tokenResponse.tokens ? 'あり' : 'なし',
      access_token: tokenResponse.tokens?.access_token ? 'あり' : 'なし'
    });
    
    oauth2Client.setCredentials(tokenResponse.tokens);
    
    console.log('✅ 認証成功、アクセストークンを設定しました');
    res.redirect('/?auth=success');
  } catch (error) {
    console.error('❌ トークン取得エラー:', error);
    console.error('エラー詳細:', error.message);
    res.redirect(`/?auth=error&message=${encodeURIComponent(error.message)}`);
  }
});

// ルートフォルダ一覧取得
app.get('/api/folders', async (req, res) => {
  try {
    console.log('📁 ルートフォルダ取得開始');
    
    // まず、すべてのフォルダを取得してみる（デバッグ用）
    console.log('🔍 デバッグ: 全フォルダ検索中...');
    const allFoldersResponse = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.folder' and trashed=false",
      pageSize: 50,
      fields: 'files(id,name,parents,modifiedTime)',
      orderBy: 'name'
    });
    
    console.log(`🔍 デバッグ: 全フォルダ数 ${allFoldersResponse.data.files?.length || 0} 個`);
    
    if (allFoldersResponse.data.files && allFoldersResponse.data.files.length > 0) {
      console.log('🔍 最初の5個のフォルダ:');
      allFoldersResponse.data.files.slice(0, 5).forEach((folder, index) => {
        console.log(`  ${index + 1}. ${folder.name} (ID: ${folder.id}, Parents: ${folder.parents || 'なし'})`);
      });
    }
    
    // ルートフォルダのみを取得
    const rootFoldersResponse = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.folder' and trashed=false and 'root' in parents",
      pageSize: 1000,
      fields: 'files(id,name,parents,modifiedTime)',
      orderBy: 'name'
    });
    
    console.log(`✅ ルートフォルダ ${rootFoldersResponse.data.files?.length || 0} 個取得`);
    
    // すべてのフォルダを返す（暫定的にデバッグのため）
    const foldersToReturn = rootFoldersResponse.data.files?.length > 0 
      ? rootFoldersResponse.data.files 
      : allFoldersResponse.data.files || [];
    
    console.log(`📤 返すフォルダ数: ${foldersToReturn.length}`);
    res.json(foldersToReturn);
  } catch (error) {
    console.error('ルートフォルダ取得エラー:', error);
    console.error('エラー詳細:', error.message);
    res.status(500).json({ error: 'フォルダの取得に失敗しました', details: error.message });
  }
});

// 指定フォルダ内のサブフォルダ取得
app.get('/api/folders/:parentId/subfolders', async (req, res) => {
  const { parentId } = req.params;
  
  try {
    console.log(`📁 サーバー: サブフォルダ取得リクエスト受信 (parentId: ${parentId})`);
    console.log(`🔍 クエリ: '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`);
    
    const response = await drive.files.list({
      q: `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      pageSize: 1000,
      fields: 'files(id,name,parents,modifiedTime)',
      orderBy: 'name'
    });
    
    console.log(`✅ サーバー: サブフォルダ ${response.data.files?.length || 0} 個取得`);
    
    if (response.data.files && response.data.files.length > 0) {
      console.log('🔍 取得したサブフォルダ:');
      response.data.files.slice(0, 3).forEach((folder, index) => {
        console.log(`  ${index + 1}. ${folder.name} (ID: ${folder.id})`);
      });
    }
    
    res.json(response.data.files || []);
  } catch (error) {
    console.error('サブフォルダ取得エラー:', error);
    console.error('エラー詳細:', error.message);
    res.status(500).json({ error: 'サブフォルダの取得に失敗しました', details: error.message });
  }
});

// フォルダの詳細情報取得（パンくずリスト用）
app.get('/api/folders/:folderId/info', async (req, res) => {
  const { folderId } = req.params;
  
  try {
    const response = await drive.files.get({
      fileId: folderId,
      fields: 'id,name,parents'
    });
    
    // 親フォルダの情報も取得してパンくずリストを作成
    const breadcrumbs = [];
    let currentFolder = response.data;
    
    // 最大10階層まで遡る（無限ループ防止）
    for (let i = 0; i < 10 && currentFolder.parents && currentFolder.parents[0] !== 'root'; i++) {
      const parentResponse = await drive.files.get({
        fileId: currentFolder.parents[0],
        fields: 'id,name,parents'
      });
      breadcrumbs.unshift(parentResponse.data);
      currentFolder = parentResponse.data;
    }
    
    res.json({
      folder: response.data,
      breadcrumbs: breadcrumbs
    });
  } catch (error) {
    console.error('フォルダ情報取得エラー:', error);
    res.status(500).json({ error: 'フォルダ情報の取得に失敗しました' });
  }
});

// 指定フォルダの写真ファイル取得
app.get('/api/folders/:folderId/photos', async (req, res) => {
  const { folderId } = req.params;
  
  try {
    const photoExtensions = ['jpg', 'jpeg', 'png', 'heic'];
    const query = `'${folderId}' in parents and trashed=false and (${photoExtensions.map(ext => `name contains '.${ext}'`).join(' or ')})`;
    
    const response = await drive.files.list({
      q: query,
      pageSize: 1000,
      fields: 'files(id,name,size,modifiedTime)',
      orderBy: 'name'
    });
    
    const files = response.data.files.filter(file => {
      const fileName = file.name.toLowerCase();
      return photoExtensions.some(ext => fileName.endsWith(`.${ext}`));
    });
    
    res.json(files);
  } catch (error) {
    console.error('写真ファイル取得エラー:', error);
    res.status(500).json({ error: '写真ファイルの取得に失敗しました' });
  }
});

// スプレッドシート一覧取得
app.get('/api/spreadsheets', async (req, res) => {
  try {
    const response = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
      pageSize: 100,
      fields: 'files(id,name,modifiedTime,owners)',
      orderBy: 'modifiedTime desc'
    });
    
    res.json(response.data.files);
  } catch (error) {
    console.error('スプレッドシート取得エラー:', error);
    res.status(500).json({ error: 'スプレッドシートの取得に失敗しました' });
  }
});

// Excelファイル処理エンドポイント
app.post('/api/process-excel', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ 
        success: false, 
        error: 'Excelファイルがアップロードされていません' 
      });
    }

    console.log(`📁 Excelファイル処理開始: ${req.file.originalname}`);
    
    // ExcelファイルをBufferから読み込み
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0]; // 最初のシートを使用
    const worksheet = workbook.Sheets[sheetName];
    
    // JSONに変換
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    // IDマッピングテーブルを作成
    const mapping = {};
    for (let i = 1; i < jsonData.length; i++) { // 1行目はヘッダーなのでスキップ
      const row = jsonData[i];
      if (row[0] && row[1]) { // ID列と名前列が両方存在する場合
        mapping[row[0].toString()] = row[1].toString();
      }
    }
    
    console.log(`📁 マッピングテーブル作成完了: ${Object.keys(mapping).length}件`);
    console.log(`📁 マッピング例:`, Object.entries(mapping).slice(0, 3));
    
    res.json({
      success: true,
      mapping: mapping,
      fileName: req.file.originalname,
      recordCount: Object.keys(mapping).length
    });
    
  } catch (error) {
    console.error('Excelファイル処理エラー:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Excelファイルの処理に失敗しました',
      details: error.message 
    });
  }
});

// スプレッドシートのヘッダー取得
app.get('/api/spreadsheets/:spreadsheetId/headers', async (req, res) => {
  const { spreadsheetId } = req.params;
  const { sheetName = 'シート1' } = req.query;
  
  try {
    console.log(`📡 ヘッダー取得リクエスト: スプレッドシートID=${spreadsheetId}, シート名=${sheetName}`);
    
    // 1行目（ヘッダー行）を取得
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!1:1`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const headers = response.data.values ? response.data.values[0] : [];
    console.log(`📡 取得したヘッダー:`, headers);
    
    res.json({
      success: true,
      headers: headers,
      sheetName: sheetName
    });
  } catch (error) {
    console.error('ヘッダー取得エラー:', error);
    res.status(500).json({ 
      success: false, 
      error: 'ヘッダーの取得に失敗しました',
      details: error.message 
    });
  }
});

// スプレッドシートに高度なデータ書き込み（動的列マッピング対応）
app.post('/api/spreadsheets/:spreadsheetId/write-advanced', async (req, res) => {
  const { spreadsheetId } = req.params;
  const { data, sheetName = 'シート1', columnMapping } = req.body;
  
  try {
    console.log(`📡 高度な書き込みリクエスト: スプレッドシートID=${spreadsheetId}`);
    console.log(`📡 書き込みデータ数: ${data ? data.length : 'undefined'}`);
    console.log(`📡 列マッピング:`, columnMapping);
    
    // 入力検証
    if (!data || !Array.isArray(data)) {
      throw new Error('データが無効です: data配列が必要です');
    }
    
    if (!columnMapping || typeof columnMapping !== 'object') {
      throw new Error('列マッピングが無効です: columnMapping オブジェクトが必要です');
    }
    
    // ヘッダー行を取得
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!1:1`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const headers = headerResponse.data.values ? headerResponse.data.values[0] : [];
    console.log(`📡 スプレッドシートのヘッダー:`, headers);
    
    // 各列の最終行を取得
    const columnLastRows = {};
    for (const [key, columnIndex] of Object.entries(columnMapping)) {
      if (columnIndex !== -1 && columnIndex < headers.length) {
        const columnLetter = String.fromCharCode(65 + columnIndex); // A, B, C...
        try {
          const columnResponse = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `${sheetName}!${columnLetter}:${columnLetter}`,
            valueRenderOption: 'UNFORMATTED_VALUE'
          });
          
          const columnValues = columnResponse.data.values || [];
          columnLastRows[key] = columnValues.length + 1; // 次の空行
          console.log(`📡 列${columnLetter}(${key})の最終行: ${columnLastRows[key]}`);
        } catch (columnError) {
          console.warn(`📡 列${columnLetter}(${key})の読み取りエラー:`, columnError.message);
          columnLastRows[key] = 2; // デフォルトで2行目から開始
        }
      } else {
        console.log(`📡 列${key}はスキップ: インデックス=${columnIndex}, ヘッダー数=${headers.length}`);
      }
    }
    
    // データを書き込み
    const updateRequests = [];
    
    for (const rowData of data) {
      for (const [key, value] of Object.entries(rowData)) {
        const columnIndex = columnMapping[key];
        if (columnIndex !== -1 && columnIndex < headers.length && columnLastRows[key]) {
          const columnLetter = String.fromCharCode(65 + columnIndex);
          const row = columnLastRows[key];
          
          // 値が undefined や null でない場合のみ追加
          if (value !== undefined && value !== null) {
            updateRequests.push({
              range: `${sheetName}!${columnLetter}${row}`,
              values: [[value]]
            });
          }
          
          columnLastRows[key]++; // 次の行に進む
        }
      }
    }
    
    console.log(`📡 書き込みリクエスト数: ${updateRequests.length}`);
    
    if (updateRequests.length > 0) {
      const batchUpdateRequest = {
        spreadsheetId,
        resource: {
          valueInputOption: 'RAW',
          data: updateRequests
        }
      };
      
      const result = await sheets.spreadsheets.values.batchUpdate(batchUpdateRequest);
      console.log(`📡 書き込み完了: ${result.data.totalUpdatedCells}セル更新`);
      
      res.json({
        success: true,
        updatedCells: result.data.totalUpdatedCells,
        updatedRows: result.data.totalUpdatedRows,
        processedRecords: data.length
      });
    } else {
      res.json({
        success: true,
        message: '書き込み可能な列が見つかりませんでした',
        processedRecords: 0
      });
    }
  } catch (error) {
    console.error('高度な書き込みエラー:', error);
    res.status(500).json({ 
      success: false, 
      error: '高度な書き込みに失敗しました',
      details: error.message 
    });
  }
});

// スプレッドシートにデータ書き込み
app.post('/api/spreadsheets/:spreadsheetId/write', async (req, res) => {
  const { spreadsheetId } = req.params;
  const { data, sheetName = 'シート1' } = req.body;
  
  try {
    // 既存データの確認
    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A2:G1000`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const existingValues = existingResponse.data.values || [];
    const nextRow = existingValues.length + 2; // 2行目から開始
    
    // データを書き込み
    const requests = data.map((row, index) => ({
      range: `${sheetName}!A${nextRow + index}:G${nextRow + index}`,
      values: [row]
    }));
    
    const batchUpdateRequest = {
      spreadsheetId,
      resource: {
        valueInputOption: 'RAW',
        data: requests
      }
    };
    
    const result = await sheets.spreadsheets.values.batchUpdate(batchUpdateRequest);
    
    res.json({
      success: true,
      updatedCells: result.data.totalUpdatedCells,
      updatedRows: result.data.totalUpdatedRows
    });
  } catch (error) {
    console.error('データ書き込みエラー:', error);
    res.status(500).json({ error: 'データの書き込みに失敗しました' });
  }
});

// サーバー起動
app.listen(PORT, () => {
  console.log(`🚀 サーバーが起動しました: http://localhost:${PORT}`);
  console.log(`📁 静的ファイルは public/ ディレクトリから提供されます`);
});