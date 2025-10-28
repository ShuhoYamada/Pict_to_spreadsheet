// セキュアなAPI管理クラス - バックエンド経由でGoogle APIsにアクセス
class SecureAPIManager {
    constructor() {
        this.selectedFolder = null;
        this.photoFiles = [];
        this.selectedSpreadsheet = null;
        this.spreadsheetData = null;
        this.currentFolderId = 'root';
        this.folderHistory = [];
        this.breadcrumbs = [];
    }

    // バックエンド経由でフォルダ一覧を取得（階層対応）
    async getFolders(parentId = 'root') {
        try {
            let url;
            if (parentId === 'root') {
                url = `${CONFIG.API_BASE_URL}/api/folders`;
                console.log('🔍 フロントエンド: ルートフォルダを取得中...');
            } else {
                url = `${CONFIG.API_BASE_URL}/api/folders/${parentId}/subfolders`;
                console.log(`🔍 フロントエンド: サブフォルダを取得中... (parentId: ${parentId})`);
            }
            
            console.log(`📡 API リクエスト: ${url}`);
            
            const response = await fetch(url, {
                credentials: 'include'
            });
            
            console.log(`📡 API レスポンス: ${response.status} ${response.statusText}`);
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('フォルダ取得エラー:', error);
            throw new Error('フォルダの取得に失敗しました: ' + error.message);
        }
    }

    // フォルダの詳細情報とパンくずリストを取得
    async getFolderInfo(folderId) {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/folders/${folderId}/info`, {
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('フォルダ情報取得エラー:', error);
            throw new Error('フォルダ情報の取得に失敗しました: ' + error.message);
        }
    }

    // バックエンド経由で指定フォルダ内の写真ファイルを取得
    async getPhotosInFolder(folderId) {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/folders/${folderId}/photos`, {
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('写真ファイル取得エラー:', error);
            throw new Error('写真ファイルの取得に失敗しました: ' + error.message);
        }
    }

    // バックエンド経由でスプレッドシート一覧を取得
    async getSpreadsheets() {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/spreadsheets`, {
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('スプレッドシート取得エラー:', error);
            throw new Error('スプレッドシートの取得に失敗しました: ' + error.message);
        }
    }

    // バックエンド経由でスプレッドシートにデータを書き込み
    async writeDataToSpreadsheet(spreadsheetId, parsedDataArray, sheetName = 'シート1') {
        try {
            const data = parsedDataArray
                .filter(item => item.isValid)
                .map(item => [
                    item.partName,
                    item.weightString,
                    item.unit,
                    item.materialId,
                    item.processId,
                    item.photoType,
                    item.notes || ''
                ]);

            const response = await fetch(`${CONFIG.API_BASE_URL}/api/spreadsheets/${spreadsheetId}/write`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    data: data,
                    sheetName: sheetName
                })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('スプレッドシート書き込みエラー:', error);
            throw new Error('データの書き込みに失敗しました: ' + error.message);
        }
    }

    // フォルダ選択機能（階層対応）
    async showFolderSelector(folderId = 'root') {
        console.log(`🔍 フロントエンド: showFolderSelector が呼び出されました (folderId: ${folderId})`);
        
        if (!authManager.isUserAuthenticated()) {
            showError('認証が必要です');
            return;
        }

        try {
            showProgress('フォルダを読み込み中...', 0);
            
            this.currentFolderId = folderId;
            console.log(`🔍 フロントエンド: getFolders を呼び出し中... (folderId: ${folderId})`);
            const folders = await this.getFolders(folderId);
            
            // パンくずリストの情報を取得（rootでない場合）
            if (folderId !== 'root') {
                const folderInfo = await this.getFolderInfo(folderId);
                this.breadcrumbs = folderInfo.breadcrumbs;
            } else {
                this.breadcrumbs = [];
            }
            
            hideProgress();

            this.renderFolderList(folders, folderId);
            document.getElementById('folder-modal').style.display = 'flex';
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // スプレッドシート選択機能
    async showSpreadsheetSelector() {
        if (!authManager.isUserAuthenticated()) {
            showError('認証が必要です');
            return;
        }

        try {
            showProgress('スプレッドシートを読み込み中...', 0);
            const spreadsheets = await this.getSpreadsheets();
            hideProgress();

            this.renderSpreadsheetList(spreadsheets);
            document.getElementById('spreadsheet-modal').style.display = 'flex';
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // フォルダリストをレンダリング（階層対応）
    renderFolderList(folders, currentFolderId = 'root') {
        const folderList = document.getElementById('folder-list');
        
        let html = '';
        
        // 操作説明を追加
        html += '<div class="folder-instructions">';
        html += '<p><strong>💡 操作方法:</strong></p>';
        html += '<ul>';
        html += '<li>📁 <strong>フォルダ名をクリック</strong> → サブフォルダを表示</li>';
        html += '<li>🎯 <strong>「このフォルダを選択」ボタン</strong> → フォルダを選択して写真を取得</li>';
        html += '</ul>';
        html += '</div>';
        
        // パンくずリストを表示
        if (currentFolderId !== 'root' || this.breadcrumbs.length > 0) {
            html += '<div class="folder-breadcrumbs">';
            html += '<button class="breadcrumb-btn" onclick="apiManager.showFolderSelector(\'root\')">🏠 ルート</button>';
            
            this.breadcrumbs.forEach(breadcrumb => {
                html += ` <span class="breadcrumb-separator">></span> `;
                html += `<button class="breadcrumb-btn" onclick="apiManager.showFolderSelector('${breadcrumb.id}')">${this.escapeHtml(breadcrumb.name)}</button>`;
            });
            
            if (currentFolderId !== 'root') {
                html += ` <span class="breadcrumb-separator">></span> <span class="current-folder">現在のフォルダ</span>`;
            }
            
            html += '</div><hr>';
        }
        
        // 戻るボタン（rootでない場合）
        if (currentFolderId !== 'root') {
            const parentId = this.breadcrumbs.length > 0 ? this.breadcrumbs[this.breadcrumbs.length - 1].id : 'root';
            html += `
                <div class="folder-item back-button" onclick="apiManager.showFolderSelector('${parentId}')">
                    <span>⬅️</span>
                    <div>
                        <strong>戻る</strong>
                        <div style="font-size: 0.8rem; color: #666;">上の階層に戻る</div>
                    </div>
                </div>
            `;
        }
        
        if (folders.length === 0) {
            html += '<p>このフォルダにはサブフォルダがありません。</p>';
        } else {
            folders.forEach(folder => {
                html += `
                    <div class="folder-item">
                        <div class="folder-content clickable-folder" onclick="apiManager.showFolderSelector('${folder.id}')" title="クリックしてサブフォルダを表示">
                            <span>📁</span>
                            <div class="folder-info">
                                <strong>${this.escapeHtml(folder.name)}</strong>
                                <div style="font-size: 0.8rem; color: #666; margin-top: 4px;">
                                    最終更新: ${new Date(folder.modifiedTime).toLocaleDateString('ja-JP')}
                                </div>
                            </div>
                            <div class="folder-expand-hint">
                                <span class="expand-icon">▶</span>
                                <div class="expand-text">サブフォルダを表示</div>
                            </div>
                        </div>
                        <div class="folder-actions">
                            <button class="select-folder-btn" onclick="apiManager.selectFolder('${folder.id}', '${this.escapeHtml(folder.name)}')" title="このフォルダの写真ファイルを取得">
                                🎯 このフォルダを選択
                            </button>
                        </div>
                    </div>
                `;
            });
        }
        
        folderList.innerHTML = html;
    }

    // スプレッドシートリストをレンダリング
    renderSpreadsheetList(spreadsheets) {
        const spreadsheetList = document.getElementById('spreadsheet-list');
        
        if (spreadsheets.length === 0) {
            spreadsheetList.innerHTML = '<p>スプレッドシートが見つかりませんでした。</p>';
            return;
        }

        spreadsheetList.innerHTML = spreadsheets.map(sheet => {
            const owner = sheet.owners && sheet.owners[0] ? sheet.owners[0].displayName : '不明';
            return `
                <div class="spreadsheet-item" onclick="apiManager.selectSpreadsheet('${sheet.id}', '${this.escapeHtml(sheet.name)}')">
                    <span>📊</span>
                    <div>
                        <strong>${this.escapeHtml(sheet.name)}</strong>
                        <div style="font-size: 0.8rem; color: #666; margin-top: 4px;">
                            所有者: ${this.escapeHtml(owner)} | 
                            最終更新: ${new Date(sheet.modifiedTime).toLocaleDateString('ja-JP')}
                        </div>
                    </div>
                </div>
            `;
        }).join('');
    }

    // フォルダを選択
    async selectFolder(folderId, folderName) {
        try {
            showProgress('写真ファイルを確認中...', 0);
            
            this.selectedFolder = {
                id: folderId,
                name: folderName
            };

            this.photoFiles = await this.getPhotosInFolder(folderId);
            
            hideProgress();
            this.updateFolderUI();
            document.getElementById('folder-modal').style.display = 'none';
            
            this.checkNextStepAvailability();
            
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // スプレッドシートを選択
    selectSpreadsheet(spreadsheetId, spreadsheetName) {
        this.selectedSpreadsheet = {
            id: spreadsheetId,
            name: spreadsheetName
        };

        this.updateSpreadsheetUI();
        document.getElementById('spreadsheet-modal').style.display = 'none';
        this.checkProcessAvailability();
    }

    // UIの更新メソッドなど（既存のコードを再利用）
    updateFolderUI() {
        const folderInfo = document.getElementById('selected-folder-info');
        const previewBox = document.getElementById('photo-files-preview');

        if (this.selectedFolder && this.photoFiles) {
            folderInfo.className = 'info-box active';
            folderInfo.innerHTML = `
                <h4>✅ 選択されたフォルダ</h4>
                <p><strong>フォルダ名:</strong> ${this.escapeHtml(this.selectedFolder.name)}</p>
                <p><strong>写真ファイル数:</strong> ${this.photoFiles.length} 個</p>
            `;

            if (this.photoFiles.length > 0) {
                previewBox.className = 'preview-box active';
                previewBox.innerHTML = `
                    <h4>📸 検出された写真ファイル (最初の10件)</h4>
                    ${this.photoFiles.slice(0, 10).map(file => this.renderFilePreview(file)).join('')}
                    ${this.photoFiles.length > 10 ? `<p><em>他 ${this.photoFiles.length - 10} 件のファイルがあります...</em></p>` : ''}
                `;
            } else {
                previewBox.className = 'preview-box active';
                previewBox.innerHTML = `
                    <h4>⚠️ 写真ファイルが見つかりません</h4>
                    <p>このフォルダには対応する写真ファイル（${CONFIG.PHOTO_EXTENSIONS.join(', ')}）が見つかりませんでした。</p>
                `;
            }
        }
    }

    updateSpreadsheetUI() {
        const spreadsheetInfo = document.getElementById('selected-spreadsheet-info');

        if (this.selectedSpreadsheet) {
            spreadsheetInfo.className = 'info-box active';
            spreadsheetInfo.innerHTML = `
                <h4>✅ 選択されたスプレッドシート</h4>
                <p><strong>名前:</strong> ${this.escapeHtml(this.selectedSpreadsheet.name)}</p>
                <p><strong>セキュリティ:</strong> バックエンドサーバー経由で安全にアクセス</p>
            `;
        }
    }

    renderFilePreview(file) {
        const parsedInfo = fileParser.parseFileName(file.name);
        const formatSize = (bytes) => {
            if (bytes === 0) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        };

        return `
            <div class="file-preview">
                <div><strong>📄 ${this.escapeHtml(file.name)}</strong></div>
                <div>サイズ: ${formatSize(file.size || 0)} | 更新: ${new Date(file.modifiedTime).toLocaleDateString('ja-JP')}</div>
                ${parsedInfo.isValid ? `
                    <div class="file-info">
                        <div class="file-info-item">部品: ${parsedInfo.partName}</div>
                        <div class="file-info-item">重量: ${parsedInfo.weight} ${parsedInfo.unit}</div>
                        <div class="file-info-item">素材: ${parsedInfo.materialId}</div>
                        <div class="file-info-item">加工: ${parsedInfo.processId}</div>
                        <div class="file-info-item">区分: ${parsedInfo.photoType}</div>
                        ${parsedInfo.notes ? `<div class="file-info-item">特記: ${parsedInfo.notes}</div>` : ''}
                    </div>
                ` : `
                    <div style="color: #e53e3e; font-size: 0.9rem;">
                        ⚠️ ファイル名の形式が正しくありません
                    </div>
                `}
            </div>
        `;
    }

    checkNextStepAvailability() {
        if (this.selectedFolder && this.photoFiles.length > 0) {
            document.getElementById('select-spreadsheet-button').disabled = false;
        }
    }

    checkProcessAvailability() {
        if (this.selectedSpreadsheet && this.photoFiles.length > 0) {
            document.getElementById('process-data-button').disabled = false;
        }
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    getSelectedPhotoFiles() {
        return this.photoFiles || [];
    }

    getSelectedFolder() {
        return this.selectedFolder;
    }

    getSelectedSpreadsheet() {
        return this.selectedSpreadsheet;
    }

    // スプレッドシートのヘッダー情報を取得（一番左のシート）
    async getSpreadsheetHeaders(spreadsheetId) {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/spreadsheets/${spreadsheetId}/headers`, {
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('スプレッドシートヘッダー取得エラー:', error);
            throw new Error('スプレッドシートのヘッダー取得に失敗しました: ' + error.message);
        }
    }

    // ハイパーリンクをスプレッドシートのセルに設定
    async setHyperlinkToCell(spreadsheetId, sheetName, cellAddress, url, displayText) {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/spreadsheets/${spreadsheetId}/set-hyperlink`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    sheetName: sheetName,
                    cellAddress: cellAddress,
                    url: url,
                    displayText: displayText
                })
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('ハイパーリンク設定エラー:', error);
            throw new Error('ハイパーリンクの設定に失敗しました: ' + error.message);
        }
    }

    // 写真ファイルの共有リンクを取得
    async getPhotoShareLink(fileId) {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/files/${fileId}/sharelink`, {
                credentials: 'include'
            });
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            return await response.json();
        } catch (error) {
            console.error('写真共有リンク取得エラー:', error);
            throw new Error('写真の共有リンク取得に失敗しました: ' + error.message);
        }
    }

    // ヘッダー列の位置を特定
    findColumnIndex(headers, targetColumns) {
        for (const targetColumn of targetColumns) {
            const index = headers.findIndex(header => 
                header && header.toString().toLowerCase().includes(targetColumn.toLowerCase())
            );
            if (index !== -1) {
                return index;
            }
        }
        return -1;
    }

    // 新仕様に基づくデータ処理と書き込み（写真区分による分岐処理とハイパーリンク設定対応）
    async processAndWriteData(spreadsheetId, photoFiles, materialMapping, processMapping) {
        try {
            // ファイル名解析
            const parser = new FileNameParser();
            const parseResults = parser.parseMultipleFiles(photoFiles);
            
            // P区分とM区分に分類
            const pTypeFiles = parseResults.valid.filter(parsed => parsed.photoType === 'p');
            const mTypeFiles = parseResults.valid.filter(parsed => parsed.isPhotoTypeM);
            
            // 番号順にソート（昇順：1,2,3,4...）
            pTypeFiles.sort((a, b) => a.number - b.number);
            mTypeFiles.sort((a, b) => a.number - b.number);
            
            console.log(`📊 解析結果: 全${photoFiles.length}件 -> P区分${pTypeFiles.length}件, M区分${mTypeFiles.length}件, 無効${parseResults.invalid.length}件`);
            console.log(`📊 P区分ファイル詳細:`, pTypeFiles.map(f => `${f.fileName} (番号: ${f.number}, 部品: ${f.partName})`));
            console.log(`📊 M区分ファイル詳細:`, mTypeFiles.map(f => `${f.fileName} (番号: ${f.number}, 部品: ${f.partName})`));
            
            if (pTypeFiles.length === 0) {
                throw new Error('処理対象のP区分ファイルがありません');
            }

            // スプレッドシートのヘッダー情報を取得
            const headerResponse = await this.getSpreadsheetHeaders(spreadsheetId);
            const headers = headerResponse.headers || [];
            const sheetName = headerResponse.sheetName || 'シート1';
            
            console.log('📊 スプレッドシートヘッダー:', headers);
            console.log('📊 対象シート名:', sheetName);
            
            if (!Array.isArray(headers) || headers.length === 0) {
                throw new Error('スプレッドシートにヘッダー行が見つかりません');
            }
            
            // 重要列の位置を特定
            const partColumnIndex = this.findColumnIndex(headers, ['構成部品', '部品', '部品名']);
            const materialColumnIndex = this.findColumnIndex(headers, ['素材', '材料']);
            
            console.log(`📊 列位置: 構成部品=${partColumnIndex}, 素材=${materialColumnIndex}`);
            
            if (partColumnIndex === -1 || materialColumnIndex === -1) {
                throw new Error('「構成部品」列または「素材」列が見つかりません');
            }
            
            // 列マッピングを生成
            const columnMapping = {
                fileName: headers.indexOf('ファイル名'),
                partName: partColumnIndex,
                weightInGrams: headers.indexOf('重量[g]'),
                weightInKilograms: headers.indexOf('重量[kg]'),
                materialId: headers.indexOf('ID'),
                processId: headers.indexOf('加工ID'),
                materialCategory: materialColumnIndex,
                materialName: headers.indexOf('項目名'),
                processName: headers.indexOf('加工方法'),
                notesText: headers.indexOf('特記事項'),
                originalUnit: headers.indexOf('元の単位')
            };
            
            console.log('📊 列マッピング:', columnMapping);
            
            // P区分ファイルのデータを変換
            const processedData = pTypeFiles.map(parsed => {
                const materialData = materialMapping ? 
                    (materialMapping[parsed.materialId] || { name: '該当なし', category: '該当なし' }) :
                    { name: '該当なし', category: '該当なし' };
                
                return {
                    fileName: parsed.fileName,
                    partName: parsed.partName,
                    weightInGrams: parsed.weightInGrams,
                    weightInKilograms: parsed.weightInKilograms,
                    materialId: parsed.materialId,
                    processId: parsed.processId,
                    notesText: parsed.notesText,
                    originalUnit: parsed.unit,
                    materialName: materialData.name,
                    materialCategory: materialData.category,
                    processName: processMapping ? processMapping[parsed.processId] || '該当なし' : '該当なし',
                    // 写真情報を追加
                    photoId: parsed.fileInfo.id,
                    photoName: parsed.fileName
                };
            });

            // P区分データをスプレッドシートに書き込み
            const response = await fetch(`${CONFIG.API_BASE_URL}/api/spreadsheets/${spreadsheetId}/write-advanced`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                credentials: 'include',
                body: JSON.stringify({
                    data: processedData,
                    columnMapping: columnMapping
                })
            });
            
            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.error || `HTTP ${response.status}: ${response.statusText}`);
            }
            
            const writeResult = await response.json();
            console.log('📊 P区分データ書き込み完了:', writeResult);
            
            // 書き込み完了後にハイパーリンクを設定
            let hyperlinkCount = 0;
            let hyperlinkErrors = [];
            
            // サーバーから返された実際の書き込み開始行を使用
            const startRow = writeResult.actualStartRow || 2;
            
            console.log(`📊 サーバーから取得した実際の書き込み開始行: ${startRow}`);
            console.log(`📊 P区分ファイル数: ${pTypeFiles.length}`);
            
            // P区分ファイルにハイパーリンクを設定（構成部品列）
            for (let i = 0; i < pTypeFiles.length; i++) {
                const pFile = pTypeFiles[i];
                const rowIndex = startRow + i;
                const partCellAddress = `${String.fromCharCode(65 + partColumnIndex)}${rowIndex}`;
                
                try {
                    console.log(`🔗 P区分ハイパーリンク設定詳細:`);
                    console.log(`   - P区分ファイル: ${pFile.fileName}`);
                    console.log(`   - インデックス: ${i}`);
                    console.log(`   - 書き込み開始行: ${startRow}`);
                    console.log(`   - 計算された行番号: ${rowIndex}`);
                    console.log(`   - 構成部品列セルアドレス: ${partCellAddress}`);
                    
                    // 写真の共有リンクを取得
                    const shareLink = await this.getPhotoShareLink(pFile.fileInfo.id);
                    
                    // 構成部品列にハイパーリンクを設定（部品名を表示テキストとして使用）
                    await this.setHyperlinkToCell(
                        spreadsheetId, 
                        sheetName, 
                        partCellAddress, 
                        shareLink.shareLink, 
                        pFile.partName  // P区分は部品名を表示テキストとして使用
                    );
                    
                    hyperlinkCount++;
                    console.log(`✅ P区分ハイパーリンク設定完了: ${partCellAddress}`);
                    
                    // API レート制限回避のため少し待機
                    if (i < pTypeFiles.length - 1) { // 最後の要素でない場合のみ待機
                        await new Promise(resolve => setTimeout(resolve, 500)); // 500ms待機
                    }
                    
                } catch (error) {
                    console.error(`❌ P区分ハイパーリンク設定エラー (${pFile.fileName}):`, error);
                    hyperlinkErrors.push({
                        fileName: pFile.fileName,
                        error: error.message,
                        type: 'P区分'
                    });
                }
            }
            
            // M区分ファイルにハイパーリンクを設定（素材列）
            let mFileIndex = 0;
            for (const mFile of mTypeFiles) {
                try {
                    console.log(`🔍 M区分ファイル処理開始: ${mFile.fileName} (番号: ${mFile.number}, 部品名: ${mFile.partName})`);
                    console.log(`🔍 利用可能なP区分ファイル:`, pTypeFiles.map(p => `${p.fileName} (番号: ${p.number}, 部品名: ${p.partName})`));
                    
                    // 対応するP区分ファイルを探す（複数の条件で検索）
                    let correspondingPFile = null;
                    
                    // 1. 番号ベースで検索
                    correspondingPFile = pTypeFiles.find(pFile => pFile.number === mFile.number);
                    
                    // 2. 番号が見つからない場合、部品名ベースで検索
                    if (!correspondingPFile) {
                        correspondingPFile = pTypeFiles.find(pFile => pFile.partName === mFile.partName);
                        console.log(`🔍 番号一致なし、部品名で検索: ${correspondingPFile ? '見つかった' : '見つからない'}`);
                    }
                    
                    // 3. 部品名も見つからない場合、ファイル名のプレフィックスで検索
                    if (!correspondingPFile) {
                        const mFilePrefix = mFile.fileName.split('_').slice(0, 2).join('_'); // 番号_部品名
                        correspondingPFile = pTypeFiles.find(pFile => {
                            const pFilePrefix = pFile.fileName.split('_').slice(0, 2).join('_');
                            return pFilePrefix === mFilePrefix;
                        });
                        console.log(`🔍 プレフィックスで検索 (${mFilePrefix}): ${correspondingPFile ? '見つかった' : '見つからない'}`);
                    }
                    
                    // 4. それでも見つからない場合、最後に処理されたP区分ファイルを使用
                    if (!correspondingPFile && pTypeFiles.length > 0) {
                        correspondingPFile = pTypeFiles[pTypeFiles.length - 1];
                        console.log(`🔍 最後のP区分ファイルを使用: ${correspondingPFile.fileName}`);
                    }
                    
                    if (!correspondingPFile) {
                        console.warn(`⚠️ M区分ファイル ${mFile.fileName} に対応するP区分ファイルが見つかりません`);
                        console.warn(`⚠️ デバッグ情報: M番号=${mFile.number}, M部品名=${mFile.partName}`);
                        console.warn(`⚠️ P区分ファイル数: ${pTypeFiles.length}`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `対応するP区分ファイルが見つかりません (番号: ${mFile.number}, 部品名: ${mFile.partName})`,
                            type: 'M区分'
                        });
                        continue;
                    }
                    
                    const pFileIndex = pTypeFiles.indexOf(correspondingPFile);
                    const rowIndex = startRow + pFileIndex;
                    const materialCellAddress = `${String.fromCharCode(65 + materialColumnIndex)}${rowIndex}`;
                    
                    console.log(`🔗 M区分ハイパーリンク設定詳細:`);
                    console.log(`   - M区分ファイル: ${mFile.fileName}`);
                    console.log(`   - 対応P区分ファイル: ${correspondingPFile.fileName}`);
                    console.log(`   - P区分インデックス: ${pFileIndex}`);
                    console.log(`   - 書き込み開始行: ${startRow}`);
                    console.log(`   - 計算された行番号: ${rowIndex}`);
                    console.log(`   - 素材列セルアドレス: ${materialCellAddress}`);
                    
                    // 写真の共有リンクを取得
                    const shareLink = await this.getPhotoShareLink(mFile.fileInfo.id);
                    
                    // セルの現在の値はサーバー側で取得するため、displayTextはnullにする
                    console.log(`🔗 M区分ハイパーリンク設定: セルの現在値をサーバー側で取得`);
                    
                    // 素材列にハイパーリンクを設定（displayTextはnull = サーバー側で現在値を取得）
                    await this.setHyperlinkToCell(
                        spreadsheetId, 
                        sheetName, 
                        materialCellAddress, 
                        shareLink.shareLink, 
                        null  // サーバー側で現在のセル値を取得
                    );
                    
                    hyperlinkCount++;
                    console.log(`✅ M区分ハイパーリンク設定完了: ${materialCellAddress}`);
                    
                    // API レート制限回避のため少し待機
                    mFileIndex++;
                    if (mFileIndex < mTypeFiles.length) { // 最後の要素でない場合のみ待機
                        await new Promise(resolve => setTimeout(resolve, 500)); // 500ms待機
                    }
                    
                } catch (error) {
                    console.error(`❌ M区分ハイパーリンク設定エラー (${mFile.fileName}):`, error);
                    hyperlinkErrors.push({
                        fileName: mFile.fileName,
                        error: error.message,
                        type: 'M区分'
                    });
                }
            }
            
            return {
                success: true,
                writtenCount: pTypeFiles.length,
                hyperlinkCount: hyperlinkCount,
                mTypeCount: mTypeFiles.length,
                invalidCount: parseResults.invalid.length,
                totalCount: photoFiles.length,
                details: writeResult,
                hyperlinkErrors: hyperlinkErrors,
                invalidFiles: parseResults.invalid.map(f => ({ fileName: f.fileName, error: f.error }))
            };

        } catch (error) {
            console.error('データ処理・書き込みエラー:', error);
            throw error;
        }
    }
}

// グローバル関数
function closeFolderModal() {
    document.getElementById('folder-modal').style.display = 'none';
}

function closeSpreadsheetModal() {
    document.getElementById('spreadsheet-modal').style.display = 'none';
}

function showProgress(message, percent) {
    const progressContainer = document.getElementById('progress-container');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    
    progressContainer.style.display = 'block';
    progressFill.style.width = percent + '%';
    progressText.textContent = message;
}

function hideProgress() {
    document.getElementById('progress-container').style.display = 'none';
}

// セキュアAPIマネージャーのインスタンス化
const apiManager = new SecureAPIManager();