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

    // エイリアス: listSpreadsheets (互換性のため)
    async listSpreadsheets() {
        return await this.getSpreadsheets();
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

    // ハイパーリンクをスプレッドシートのセルに設定（エクスポネンシャル・バックオフ付き）
    async setHyperlinkToCell(spreadsheetId, sheetName, cellAddress, url, displayText) {
        const maxRetries = 10;
        let retryCount = 0;
        
        while (retryCount < maxRetries) {
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
                
                if (response.ok) {
                    // 成功した場合
                    const result = await response.json();
                    if (retryCount > 0) {
                        console.log(`✅ ハイパーリンク設定成功 (${retryCount + 1}回目の試行): ${cellAddress}`);
                    }
                    return result;
                }
                
                // 500エラーの場合はリトライ対象
                if (response.status === 500) {
                    throw new Error(`HTTP 500: Internal Server Error`);
                }
                
                // その他のエラーはリトライしない
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                
            } catch (error) {
                retryCount++;
                
                // 500エラー以外、または最大リトライ回数に達した場合
                if (!error.message.includes('HTTP 500') || retryCount >= maxRetries) {
                    console.error(`❌ ハイパーリンク設定エラー (最終失敗, ${retryCount}回試行): ${cellAddress}`, error);
                    throw new Error('ハイパーリンクの設定に失敗しました: ' + error.message);
                }
                
                // エクスポネンシャル・バックオフ：1秒, 2秒, 4秒, 8秒, 16秒, 32秒, 7回目以降は32秒固定
                const waitTime = retryCount <= 6 ? Math.pow(2, retryCount - 1) * 1000 : 32000;
                console.warn(`⚠️ ハイパーリンク設定失敗 (${retryCount}/${maxRetries}回目): ${cellAddress} - ${waitTime/1000}秒後にリトライ`);
                
                await new Promise(resolve => setTimeout(resolve, waitTime));
            }
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
    async processAndWriteData(spreadsheetId, photoFiles, materialMapping, processMapping, implementerMapping) {
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
                implementerName: headers.indexOf('実施者名'),
                implementationDate: headers.indexOf('実施日'),
                notesText: headers.indexOf('特記事項'),
                originalUnit: headers.indexOf('元の単位')
            };
            
            console.log('📊 列マッピング:', columnMapping);
            
            // 現在の年月を取得（yyyy/mmフォーマット）
            const now = new Date();
            const year = now.getFullYear();
            const month = String(now.getMonth() + 1).padStart(2, '0'); // 月は0始まりなので+1、2桁にゼロパディング
            const implementationDate = `${year}/${month}`;
            
            console.log(`📅 実施日: ${implementationDate}`);
            
            // P区分ファイルのデータを変換
            const processedData = pTypeFiles.map(parsed => {
                const materialData = materialMapping ? 
                    (materialMapping[parsed.materialId] || { name: '該当なし', category: '該当なし' }) :
                    { name: '該当なし', category: '該当なし' };
                
                const implementerName = implementerMapping ? 
                    (implementerMapping[parsed.implementerId] || '該当なし') : '該当なし';
                
                return {
                    fileName: parsed.fileName,
                    partName: parsed.partName,
                    weightInGrams: parsed.weightInGrams,
                    weightInKilograms: parsed.weightInKilograms,
                    materialId: parsed.materialId,
                    processId: parsed.processId,
                    implementerName: implementerName,
                    implementationDate: implementationDate,
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
            const startRow = writeResult.startRow || writeResult.actualStartRow || 2;
            
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
                    console.log(`🔍 M区分ペア処理開始: ${mFile.fileName} (${mFileIndex + 1}/${mTypeFiles.length})`);
                    console.log(`🔍 M区分詳細: ペア番号=${mFile.number}, 部品名=${mFile.partName}`);
                    console.log(`🔍 探索対象P区分番号:`, pTypeFiles.map(p => p.number).sort((a, b) => a - b));
                    
                    // 対応するP区分ファイルを探す（ペア番号ベースで検索）
                    let correspondingPFile = null;
                    
                    // ペア形式（1,1,2,2,3,3...）では番号で完全一致検索
                    correspondingPFile = pTypeFiles.find(pFile => pFile.number === mFile.number);
                    console.log(`🔍 ペア番号で検索 (番号: ${mFile.number}): ${correspondingPFile ? '見つかった' : '見つからない'}`);
                    
                    if (correspondingPFile) {
                        console.log(`✅ ペア確認: M区分 ${mFile.fileName} ↔ P区分 ${correspondingPFile.fileName}`);
                    }
                    
                    if (!correspondingPFile) {
                        console.error(`❌ M区分ファイル ${mFile.fileName} のペアP区分ファイルが見つかりません`);
                        console.error(`❌ ペア番号: ${mFile.number}`);
                        console.error(`❌ 利用可能なP区分番号:`, pTypeFiles.map(p => p.number).sort((a, b) => a - b));
                        console.error(`❌ 期待されるペアファイル名: ${mFile.number}_*_P_*.jpg`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `ペアのP区分ファイルが見つかりません (番号: ${mFile.number})`,
                            type: 'M区分ペアエラー'
                        });
                        continue;
                    }
                    
                    // pFileIndexを取得
                    const pFileIndex = pTypeFiles.indexOf(correspondingPFile);
                    
                    // pFileIndexの妥当性チェック
                    if (pFileIndex < 0 || pFileIndex >= pTypeFiles.length) {
                        console.error(`❌ P区分ファイルインデックスが不正: ${pFileIndex} (範囲: 0-${pTypeFiles.length - 1})`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `P区分ファイルインデックスが不正です (${pFileIndex})`,
                            type: 'M区分'
                        });
                        continue;
                    }
                    
                    const rowIndex = startRow + pFileIndex;
                    
                    // 素材列インデックスの検証
                    if (materialColumnIndex === -1 || materialColumnIndex === undefined || materialColumnIndex === null) {
                        console.error(`❌ 素材列インデックスが無効: ${materialColumnIndex}`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `素材列が見つかりません (インデックス: ${materialColumnIndex})`,
                            type: 'M区分'
                        });
                        continue;
                    }
                    
                    if (typeof materialColumnIndex !== 'number' || materialColumnIndex < 0) {
                        console.error(`❌ 素材列インデックスが数値でないか負の値: ${materialColumnIndex} (型: ${typeof materialColumnIndex})`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `素材列インデックスが不正です (${materialColumnIndex})`,
                            type: 'M区分'
                        });
                        continue;
                    }
                    
                    const materialCellAddress = `${String.fromCharCode(65 + materialColumnIndex)}${rowIndex}`;
                    
                    console.log(`🔗 ペアハイパーリンク設定詳細:`);
                    console.log(`   - Mファイル: ${mFile.fileName} (ペア番号: ${mFile.number})`);
                    console.log(`   - ペアPファイル: ${correspondingPFile.fileName} (番号: ${correspondingPFile.number})`);
                    console.log(`   - P区分のスプレッドシート行インデックス: ${pFileIndex}`);
                    console.log(`   - P区分のスプレッドシート行番号: ${rowIndex}`);
                    console.log(`   - 素材列インデックス: ${materialColumnIndex} (型: ${typeof materialColumnIndex})`);
                    console.log(`   - M写真リンク設定先セル: ${materialCellAddress} (P区分行の素材列)`);
                    
                    // 行番号の妥当性チェック
                    if (rowIndex <= 0 || rowIndex > 1000) {
                        console.error(`❌ 計算された行番号が範囲外: ${rowIndex} (1-1000の範囲を期待)`);
                        hyperlinkErrors.push({
                            fileName: mFile.fileName,
                            error: `計算された行番号が範囲外です (${rowIndex})`,
                            type: 'M区分'
                        });
                        continue;
                    }
                    
                    // 写真の共有リンクを取得
                    const shareLink = await this.getPhotoShareLink(mFile.fileInfo.id);
                    
                    // 対応するP区分ファイルの素材IDから素材名を取得（スプレッドシートに書き込まれた値と同じ）
                    const materialData = materialMapping ? 
                        (materialMapping[correspondingPFile.materialId] || { name: '該当なし', category: '該当なし' }) :
                        { name: '該当なし', category: '該当なし' };
                    const materialName = materialData.name;
                    
                    console.log(`🔗 M区分ハイパーリンク設定: P区分(${correspondingPFile.materialId})の素材名="${materialName}" を表示テキストとして使用`);
                    
                    // 素材列にハイパーリンクを設定（素材名を明示的に指定）
                    await this.setHyperlinkToCell(
                        spreadsheetId, 
                        sheetName, 
                        materialCellAddress, 
                        shareLink.shareLink, 
                        materialName  // P区分の素材名を表示テキストとして使用
                    );
                    
                    hyperlinkCount++;
                    console.log(`✅ ペアハイパーリンク設定完了: M区分(${mFile.number}) → P区分行の素材列(${materialCellAddress})`);
                    
                    mFileIndex++;
                    
                } catch (error) {
                    console.error(`❌ ペアハイパーリンク設定エラー (${mFile.fileName}, ペア番号: ${mFile.number}):`, error);
                    hyperlinkErrors.push({
                        fileName: mFile.fileName,
                        error: error.message,
                        type: 'M区分ペア処理',
                        pairNumber: mFile.number
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