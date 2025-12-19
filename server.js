const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
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
    console.log('📁 ルートレベルフォルダ取得開始');
    
    // マイドライブのルートフォルダを取得
    console.log('🔍 マイドライブのルートフォルダを検索中...');
    const myDriveResponse = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.folder' and trashed=false and 'root' in parents",
      pageSize: 1000,
      fields: 'files(id,name,parents,modifiedTime)',
      orderBy: 'name'
    });
    
    console.log(`✅ マイドライブのルートフォルダ ${myDriveResponse.data.files?.length || 0} 個取得`);
    
    // 共有ドライブのフォルダを取得
    console.log('🔍 共有ドライブのフォルダを検索中...');
    let sharedDriveFolders = [];
    try {
      const sharedDrivesResponse = await drive.drives.list({
        pageSize: 100,
        fields: 'drives(id,name)'
      });
      
      if (sharedDrivesResponse.data.drives && sharedDrivesResponse.data.drives.length > 0) {
        console.log(`🔍 共有ドライブ ${sharedDrivesResponse.data.drives.length} 個発見`);
        
        // 各共有ドライブをフォルダとして扱う
        sharedDriveFolders = sharedDrivesResponse.data.drives.map(drive => ({
          id: drive.id,
          name: `📂 ${drive.name} (共有ドライブ)`,
          modifiedTime: new Date().toISOString(),
          isSharedDrive: true
        }));
      }
    } catch (sharedDriveError) {
      console.log('ℹ️ 共有ドライブへのアクセス権限がないか、利用できません');
    }
    
    // 自分と共有されているフォルダ（親フォルダがないもの）を取得
    console.log('🔍 共有フォルダを検索中...');
    const sharedFoldersResponse = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.folder' and trashed=false and sharedWithMe=true",
      pageSize: 1000,
      fields: 'files(id,name,parents,modifiedTime,shared,owners)',
      orderBy: 'name'
    });
    
    // 親フォルダを持たない共有フォルダのみをトップレベルに表示
    const topLevelSharedFolders = sharedFoldersResponse.data.files?.filter(folder => 
      !folder.parents || folder.parents.length === 0
    ) || [];
    
    console.log(`✅ トップレベル共有フォルダ ${topLevelSharedFolders.length} 個取得`);
    
    // 共有フォルダに所有者情報を追加
    const sharedFoldersWithOwner = topLevelSharedFolders.map(folder => ({
      ...folder,
      name: `🔗 ${folder.name} (共有)`,
      ownerName: folder.owners && folder.owners[0] ? folder.owners[0].displayName : '不明'
    }));
    
    // すべてのフォルダを統合
    const allFolders = [
      ...myDriveResponse.data.files || [],
      ...sharedDriveFolders,
      ...sharedFoldersWithOwner
    ];
    
    console.log(`📤 返すフォルダ数: ${allFolders.length} (マイドライブ: ${myDriveResponse.data.files?.length || 0}, 共有ドライブ: ${sharedDriveFolders.length}, 共有フォルダ: ${sharedFoldersWithOwner.length})`);
    
    res.json(allFolders);
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
    
    let subfolders = [];
    
    // 共有ドライブの場合は特別な処理
    // 共有ドライブIDかどうかを判定（通常のフォルダIDと形式が異なる場合がある）
    let isSharedDrive = false;
    
    try {
      // まず通常のフォルダとして処理
      console.log(`🔍 クエリ: '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`);
      
      const response = await drive.files.list({
        q: `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        pageSize: 1000,
        fields: 'files(id,name,parents,modifiedTime,driveId)',
        orderBy: 'name',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
      
      subfolders = response.data.files || [];
      console.log(`✅ サーバー: サブフォルダ ${subfolders.length} 個取得`);
      
    } catch (firstError) {
      console.log('ℹ️ 通常のフォルダとして取得できませんでした。共有ドライブとして再試行します。');
      
      // 共有ドライブのルートフォルダを取得
      try {
        const sharedDriveResponse = await drive.files.list({
          q: `mimeType='application/vnd.google-apps.folder' and trashed=false`,
          pageSize: 1000,
          fields: 'files(id,name,parents,modifiedTime)',
          orderBy: 'name',
          driveId: parentId,
          corpora: 'drive',
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        });
        
        // ルートレベルのフォルダのみをフィルタ（parentsがない、またはdriveIdのみを持つもの）
        subfolders = sharedDriveResponse.data.files?.filter(folder => 
          !folder.parents || folder.parents.length === 0
        ) || [];
        
        console.log(`✅ 共有ドライブのルートフォルダ ${subfolders.length} 個取得`);
        isSharedDrive = true;
        
      } catch (sharedDriveError) {
        console.error('❌ 共有ドライブとしても取得失敗:', sharedDriveError.message);
        throw firstError; // 元のエラーをスロー
      }
    }
    
    if (subfolders.length > 0) {
      console.log('🔍 取得したサブフォルダ:');
      subfolders.slice(0, 3).forEach((folder, index) => {
        console.log(`  ${index + 1}. ${folder.name} (ID: ${folder.id})`);
      });
    }
    
    res.json(subfolders);
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
      fields: 'id,name,parents,driveId',
      supportsAllDrives: true
    });
    
    // 親フォルダの情報も取得してパンくずリストを作成
    const breadcrumbs = [];
    let currentFolder = response.data;
    
    // 最大10階層まで遡る（無限ループ防止）
    for (let i = 0; i < 10 && currentFolder.parents && currentFolder.parents[0] !== 'root'; i++) {
      try {
        const parentResponse = await drive.files.get({
          fileId: currentFolder.parents[0],
          fields: 'id,name,parents',
          supportsAllDrives: true
        });
        breadcrumbs.unshift(parentResponse.data);
        currentFolder = parentResponse.data;
      } catch (error) {
        console.log('親フォルダ取得エラー（継続）:', error.message);
        break; // エラーが発生したら終了
      }
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
    console.log(`📷 写真ファイル取得開始 (folderId: ${folderId})`);
    
    const photoExtensions = ['jpg', 'jpeg', 'png', 'heic'];
    const query = `'${folderId}' in parents and trashed=false and (${photoExtensions.map(ext => `name contains '.${ext}'`).join(' or ')})`;
    
    console.log(`🔍 クエリ: ${query}`);
    
    const response = await drive.files.list({
      q: query,
      pageSize: 1000,
      fields: 'files(id,name,size,modifiedTime,webViewLink)',
      orderBy: 'name',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    
    const files = response.data.files.filter(file => {
      const fileName = file.name.toLowerCase();
      return photoExtensions.some(ext => fileName.endsWith(`.${ext}`));
    });
    
    console.log(`✅ 写真ファイル ${files.length} 個取得`);
    
    res.json(files);
  } catch (error) {
    console.error('写真ファイル取得エラー:', error);
    console.error('エラー詳細:', error.message);
    res.status(500).json({ error: '写真ファイルの取得に失敗しました', details: error.message });
  }
});

// 写真ファイルの共有リンク取得（個別）
app.get('/api/files/:fileId/sharelink', async (req, res) => {
  const { fileId } = req.params;
  
  try {
    console.log(`🔗 共有リンク取得開始 (fileId: ${fileId})`);
    
    // リトライメカニズム付きでファイル情報を取得
    let file;
    let retryCount = 0;
    const maxRetries = 3;
    
    while (retryCount < maxRetries) {
      try {
        const response = await drive.files.get({
          fileId: fileId,
          fields: 'id,name,webViewLink,permissions'
        });
        file = response.data;
        console.log(`✅ ファイル情報取得成功 (試行回数: ${retryCount + 1}): ${file.name}, webViewLink: ${file.webViewLink || 'なし'}`);
        break;
      } catch (retryError) {
        retryCount++;
        console.warn(`⚠️ ファイル情報取得失敗 (試行回数: ${retryCount}/${maxRetries}):`, {
          fileId,
          error: retryError.message,
          errorCode: retryError.code
        });
        
        if (retryCount >= maxRetries) {
          throw retryError;
        }
        
        // リトライ前に少し待機
        const waitTime = Math.pow(2, retryCount) * 500; // 1秒, 2秒, 4秒
        console.log(`⏳ ${waitTime/1000}秒待機してリトライします...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    }
    
    // webViewLinkが存在しない場合、ファイルを公開共有に設定
    if (!file.webViewLink) {
      console.log('🔗 webViewLinkが見つからないため、共有設定を確認中...');
      
      try {
        // ファイルを「リンクを知っている全員が閲覧可能」に設定
        await drive.permissions.create({
          fileId: fileId,
          requestBody: {
            role: 'reader',
            type: 'anyone'
          }
        });
        
        console.log('✅ 共有設定完了: リンクを知っている全員が閲覧可能');
        
        // 再度ファイル情報を取得
        const updatedResponse = await drive.files.get({
          fileId: fileId,
          fields: 'id,name,webViewLink'
        });
        
        const updatedFile = updatedResponse.data;
        console.log(`🔗 更新後のwebViewLink: ${updatedFile.webViewLink}`);
        
        res.json({
          success: true,
          fileId: fileId,
          fileName: updatedFile.name,
          shareLink: updatedFile.webViewLink,
          retryCount: retryCount
        });
      } catch (permissionError) {
        console.error('❌ 共有設定エラー:', {
          fileId,
          error: permissionError.message,
          errorCode: permissionError.code,
          errorDetails: permissionError.response?.data || 'レスポンスデータなし'
        });
        res.status(500).json({ 
          error: 'ファイルの共有設定に失敗しました', 
          details: permissionError.message,
          fileId: fileId,
          errorCode: permissionError.code
        });
      }
    } else {
      res.json({
        success: true,
        fileId: fileId,
        fileName: file.name,
        shareLink: file.webViewLink,
        retryCount: retryCount
      });
    }
    
  } catch (error) {
    console.error('❌ 共有リンク取得エラー (最終失敗):', {
      fileId,
      error: error.message,
      errorCode: error.code,
      stack: error.stack,
      errorDetails: error.response?.data || 'レスポンスデータなし'
    });
    res.status(500).json({ 
      error: '共有リンクの取得に失敗しました', 
      details: error.message,
      fileId: fileId,
      errorCode: error.code,
      errorDetails: error.response?.data || 'レスポンスデータなし'
    });
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
app.post('/api/process-excel', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ 
        success: false, 
        error: 'Excelファイルがアップロードされていません' 
      });
    }

    console.log(`📁 Excelファイル処理開始: ${req.file.originalname}`);
    
    // ExcelJSでExcelファイルをBufferから読み込み
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    
    const worksheet = workbook.worksheets[0]; // 最初のシートを使用
    
    // JSONに変換
    const jsonData = [];
    worksheet.eachRow((row, rowNumber) => {
      const rowData = [];
      row.eachCell((cell, colNumber) => {
        rowData[colNumber - 1] = cell.value;
      });
      jsonData.push(rowData);
    });
    
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
  
  try {
    console.log(`📡 ヘッダー取得リクエスト: スプレッドシートID=${spreadsheetId}`);
    
    // スプレッドシートのメタデータを取得して一番左のシート名を取得
    const metadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties.title,sheets.properties.sheetId'
    });
    
    const firstSheetName = metadataResponse.data.sheets[0].properties.title;
    console.log(`📡 一番左のシート名: ${firstSheetName}`);
    
    // 1行目（ヘッダー行）を取得
    const headerResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${firstSheetName}!1:1`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const headers = headerResponse.data.values ? headerResponse.data.values[0] : [];
    console.log(`📡 取得したヘッダー:`, headers);
    
    // データ行数を取得（現在のデータ範囲を特定）
    let dataRowCount = 0;
    try {
      const dataResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${firstSheetName}!A:A`,
        valueRenderOption: 'UNFORMATTED_VALUE'
      });
      dataRowCount = (dataResponse.data.values || []).length;
      console.log(`📡 データ行数: ${dataRowCount}`);
    } catch (error) {
      console.warn('データ行数取得時のエラー:', error.message);
    }
    
    res.json({
      success: true,
      headers: headers,
      sheetName: firstSheetName,
      dataRowCount: dataRowCount
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
  const { data, columnMapping } = req.body;
  
  try {
    console.log(`📡 高度な書き込みリクエスト: スプレッドシートID=${spreadsheetId}`);
    console.log(`📡 書き込みデータ数: ${data ? data.length : 'undefined'}`);
    console.log(`📡 列マッピング:`, columnMapping);
    
    // スプレッドシートのメタデータを取得して一番左のシート名を取得
    const metadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties.title,sheets.properties.sheetId'
    });
    
    const firstSheetName = metadataResponse.data.sheets[0].properties.title;
    console.log(`📡 書き込み対象シート: ${firstSheetName}`);
    
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
      range: `${firstSheetName}!1:1`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const headers = headerResponse.data.values ? headerResponse.data.values[0] : [];
    console.log(`📡 スプレッドシートのヘッダー:`, headers);
    
    // 構成部品列の次の空行を取得（これを基準行とする）
    const partColumnIndex = columnMapping['partName'];
    let startRow = 2; // デフォルトは2行目
    
    if (partColumnIndex !== -1 && partColumnIndex < headers.length) {
      const partColumnLetter = String.fromCharCode(65 + partColumnIndex);
      try {
        const partColumnResponse = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${firstSheetName}!${partColumnLetter}:${partColumnLetter}`,
          valueRenderOption: 'UNFORMATTED_VALUE'
        });
        
        const partColumnValues = partColumnResponse.data.values || [];
        startRow = partColumnValues.length + 1; // 次の空行
        console.log(`📡 構成部品列(${partColumnLetter})の次の空行: ${startRow}`);
      } catch (error) {
        console.warn(`📡 構成部品列の読み取りエラー:`, error.message);
        startRow = 2;
      }
    } else {
      console.warn(`📡 構成部品列が見つかりません。デフォルト行(2)から開始します。`);
    }
    
    console.log(`📡 書き込み開始行（構成部品基準）: ${startRow}`);
    
    // データを書き込み（全列を構成部品と同じ行に書き込む）
    const updateRequests = [];
    let currentRow = startRow; // 構成部品列の空行から開始
    
    for (const rowData of data) {
      for (const [key, value] of Object.entries(rowData)) {
        const columnIndex = columnMapping[key];
        if (columnIndex !== -1 && columnIndex < headers.length) {
          const columnLetter = String.fromCharCode(65 + columnIndex);
          
          // 値が undefined や null でない場合のみ追加
          if (value !== undefined && value !== null) {
            updateRequests.push({
              range: `${firstSheetName}!${columnLetter}${currentRow}`,
              values: [[value]]
            });
          }
        }
      }
      currentRow++; // 次のデータ行に進む
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
        processedRecords: data.length,
        startRow: startRow // 構成部品基準の書き込み開始行
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
  const { data } = req.body;
  
  try {
    // スプレッドシートのメタデータを取得して一番左のシート名を取得
    const metadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties.title,sheets.properties.sheetId'
    });
    
    const firstSheetName = metadataResponse.data.sheets[0].properties.title;
    console.log(`📡 書き込み対象シート: ${firstSheetName}`);
    
    // 既存データの確認
    const existingResponse = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${firstSheetName}!A2:G1000`,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    
    const existingValues = existingResponse.data.values || [];
    const nextRow = existingValues.length + 2; // 2行目から開始
    
    // データを書き込み
    const requests = data.map((row, index) => ({
      range: `${firstSheetName}!A${nextRow + index}:G${nextRow + index}`,
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

// ハイパーリンクをスプレッドシートのセルに設定
app.post('/api/spreadsheets/:spreadsheetId/set-hyperlink', async (req, res) => {
  const { spreadsheetId } = req.params;
  const { sheetName, cellAddress, url, displayText } = req.body;
  
  try {
    console.log(`🔗 ハイパーリンク設定開始: ${cellAddress} -> ${url}`);
    console.log(`🔗 リクエストボディ詳細:`, req.body);
    console.log(`🔗 リクエスト元IP: ${req.ip}`);
    
    // 入力検証
    if (!cellAddress || !url) {
      return res.status(400).json({ 
        error: 'セルアドレスとURLは必須です',
        received: { cellAddress, url }
      });
    }
    
    // セルアドレスをパース (例: "A5" -> row=5, col=1)
    const match = cellAddress.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      return res.status(400).json({ 
        error: '無効なセルアドレス形式です。例: A5, B10',
        received: cellAddress
      });
    }
    
    const column = match[1];
    const row = parseInt(match[2]);
    
    console.log(`🔗 セルアドレス解析: 列=${column}, 行=${row}`);
    
    // 列文字を数値に変換 (A=0, B=1, ...)
    let columnIndex = 0;
    for (let i = 0; i < column.length; i++) {
      columnIndex = columnIndex * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    columnIndex--; // 0ベースのインデックスに調整
    
    // シート名が指定されていない場合、最初のシートを使用
    let targetSheetName = sheetName;
    if (!targetSheetName) {
      const initialMetadataResponse = await sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties.title'
      });
      targetSheetName = initialMetadataResponse.data.sheets[0].properties.title;
    }
    
    console.log(`🔗 対象シート名: ${targetSheetName}`);
    
    // シートIDを取得
    const sheetMetadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties'
    });
    
    const hyperlinkTargetSheet = sheetMetadataResponse.data.sheets.find(sheet => 
      sheet.properties.title === targetSheetName
    );
    
    if (!hyperlinkTargetSheet) {
      return res.status(404).json({ 
        error: `シート '${targetSheetName}' が見つかりません`,
        availableSheets: sheetMetadataResponse.data.sheets.map(s => s.properties.title)
      });
    }
    
    // シートの行数と列数を取得してセルアドレスが有効かチェック
    const sheetProperties = hyperlinkTargetSheet.properties;
    const maxRows = sheetProperties.gridProperties?.rowCount || 1000;
    const maxCols = sheetProperties.gridProperties?.columnCount || 26;
    
    console.log(`🔗 シート情報: 最大行数=${maxRows}, 最大列数=${maxCols}, 指定行=${row}, 指定列インデックス=${columnIndex + 1}`);
    
    // 行と列の範囲チェック
    if (row > maxRows) {
      return res.status(400).json({ 
        error: `指定された行 ${row} がシートの最大行数 ${maxRows} を超えています`,
        cellAddress,
        sheetName: targetSheetName,
        sheetInfo: { maxRows, maxCols }
      });
    }
    
    if (columnIndex + 1 > maxCols) {
      return res.status(400).json({ 
        error: `指定された列 ${column} (インデックス ${columnIndex + 1}) がシートの最大列数 ${maxCols} を超えています`,
        cellAddress,
        sheetName: targetSheetName,
        sheetInfo: { maxRows, maxCols }
      });
    }
    
    // 現在のセルの値を取得
    let currentCellValue = displayText || '';
    try {
      const currentValueResponse = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${targetSheetName}!${cellAddress}`,
        valueRenderOption: 'UNFORMATTED_VALUE'
      });
      
      if (currentValueResponse.data.values && currentValueResponse.data.values[0] && currentValueResponse.data.values[0][0]) {
        currentCellValue = currentValueResponse.data.values[0][0].toString();
        console.log(`🔗 現在のセル値を取得: ${currentCellValue}`);
      } else {
        console.log(`🔗 セルは空または値なし、displayTextを使用: ${currentCellValue}`);
      }
    } catch (error) {
      console.warn(`🔗 現在のセル値取得エラー、displayTextを使用: ${error.message}`);
      // エラーの場合はdisplayTextを使用（フォールバック）
    }
    
    // URLとセル値をエスケープして安全にする
    const escapedUrl = url.replace(/"/g, '""'); // ダブルクォートをエスケープ
    const escapedCellValue = currentCellValue.replace(/"/g, '""'); // ダブルクォートをエスケープ
    
    console.log(`🔗 エスケープ後 URL: ${escapedUrl}`);
    console.log(`🔗 エスケープ後 セル値: ${escapedCellValue}`);
    
    // ハイパーリンクを設定するためのbatchUpdateリクエスト
    const requests = [{
      updateCells: {
        range: {
          sheetId: hyperlinkTargetSheet.properties.sheetId,
          startRowIndex: row - 1,
          endRowIndex: row,
          startColumnIndex: columnIndex,
          endColumnIndex: columnIndex + 1
        },
        rows: [{
          values: [{
            userEnteredValue: {
              formulaValue: `=HYPERLINK("${escapedUrl}", "${escapedCellValue}")`
            }
          }]
        }],
        fields: 'userEnteredValue'
      }
    }];
    
    // シートIDを取得
    const metadataResponse = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties'
    });
    
    const targetSheet = metadataResponse.data.sheets.find(sheet => 
      sheet.properties.title === targetSheetName
    );
    
    if (!targetSheet) {
      return res.status(404).json({ 
        error: `シート '${targetSheetName}' が見つかりません`,
        availableSheets: metadataResponse.data.sheets.map(s => s.properties.title)
      });
    }
    
    // 正しいシートIDを設定
    requests[0].updateCells.range.sheetId = targetSheet.properties.sheetId;
    
    const batchUpdateRequest = {
      spreadsheetId,
      resource: {
        requests: requests
      }
    };
    
    console.log(`🔗 ハイパーリンク設定リクエスト詳細:`, {
      cellAddress,
      url: escapedUrl,
      displayText: escapedCellValue,
      sheetName: targetSheetName,
      sheetId: hyperlinkTargetSheet.properties.sheetId
    });
    
    // リトライメカニズム付きでハイパーリンク設定を実行
    let result;
    let retryCount = 0;
    const maxRetries = 5;
    
    while (retryCount < maxRetries) {
      try {
        result = await sheets.spreadsheets.batchUpdate(batchUpdateRequest);
        console.log(`✅ ハイパーリンク設定成功 (試行回数: ${retryCount + 1}): ${cellAddress} -> ${url}`);
        break;
      } catch (retryError) {
        retryCount++;
        console.warn(`⚠️ ハイパーリンク設定失敗 (試行回数: ${retryCount}/${maxRetries}):`, {
          cellAddress,
          error: retryError.message,
          errorCode: retryError.code,
          errorDetails: retryError.response?.data || 'レスポンスデータなし'
        });
        
        if (retryCount >= maxRetries) {
          throw retryError;
        }
        
        // エクスポネンシャル・バックオフ：1秒, 2秒, 4秒, 8秒, 16秒
        const waitTime = Math.pow(2, retryCount - 1) * 1000;
        console.log(`⏳ ${waitTime/1000}秒待機してリトライします... (エクスポネンシャル・バックオフ)`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    }
    
    res.json({
      success: true,
      cellAddress: cellAddress,
      url: url,
      displayText: escapedCellValue,
      sheetName: targetSheetName,
      retryCount: retryCount,
      result: result.data
    });
    
  } catch (error) {
    console.error('❌ ハイパーリンク設定エラー (最終失敗):', {
      cellAddress,
      url,
      error: error.message,
      errorCode: error.code,
      stack: error.stack,
      errorDetails: error.response?.data || 'レスポンスデータなし'
    });
    
    res.status(500).json({ 
      error: 'ハイパーリンクの設定に失敗しました', 
      details: error.message,
      cellAddress: cellAddress,
      url: url,
      errorCode: error.code,
      errorDetails: error.response?.data || 'レスポンスデータなし'
    });
  }
});

// サーバー起動
app.listen(PORT, () => {
  console.log(`🚀 サーバーが起動しました: http://localhost:${PORT}`);
  console.log(`📁 静的ファイルは public/ ディレクトリから提供されます`);
});