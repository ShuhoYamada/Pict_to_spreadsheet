// ID対応表管理クラス - Google Sheets API対応版
class MappingTableManager {
    constructor() {
        this.materialMapping = null;
        this.processMapping = null;
        this.implementerMapping = null;
        this.idMasterSpreadsheetId = null;
        this.idMasterSpreadsheetName = null;
        this.eventListenersSetup = false; // 重複防止フラグ
        this.currentFolderId = 'root'; // 現在のフォルダID
        this.breadcrumbs = []; // パンくずリスト
    }

    // 初期化
    initialize() {
        this.setupEventListeners();
    }

    // イベントリスナーの設定
    setupEventListeners() {
        if (this.eventListenersSetup) {
            console.log('⚠️ イベントリスナーは既に設定済みです - 重複登録を防止');
            return;
        }

        console.log('🔧 マッピングマネージャーのイベントリスナーを設定中...');

        // IDマスター選択ボタン
        const idMasterButton = document.getElementById('select-id-master-button');
        if (idMasterButton) {
            idMasterButton.addEventListener('click', async () => {
                await this.showIdMasterModal();
            });
        }

        this.eventListenersSetup = true;
        console.log('✅ マッピングマネージャーのイベントリスナー設定完了');
    }

    // IDマスター選択モーダルを表示（階層対応）
    async showIdMasterModal(folderId = 'root') {
        try {
            showProgress('フォルダとスプレッドシートを読み込み中...', 0);
            
            this.currentFolderId = folderId;
            
            // フォルダとスプレッドシートを取得
            const folders = await apiManager.getFolders(folderId);
            const spreadsheets = await this.getSpreadsheetsByFolder(folderId);
            
            // パンくずリストの情報を取得（rootでない場合）
            if (folderId !== 'root') {
                const folderInfo = await apiManager.getFolderInfo(folderId);
                this.breadcrumbs = folderInfo.breadcrumbs;
            } else {
                this.breadcrumbs = [];
            }
            
            hideProgress();
            
            this.renderIdMasterList(folders, spreadsheets, folderId);
            document.getElementById('id-master-modal').style.display = 'flex';
            
        } catch (error) {
            hideProgress();
            showError('フォルダとスプレッドシートの取得に失敗しました: ' + error.message);
            console.error('フォルダ・スプレッドシート取得エラー:', error);
        }
    }

    // フォルダ内のスプレッドシートを取得
    async getSpreadsheetsByFolder(folderId) {
        try {
            const response = await fetch(
                `${CONFIG.API_BASE_URL}/api/folders/${folderId}/spreadsheets`,
                {
                    method: 'GET',
                    credentials: 'include'
                }
            );

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'スプレッドシートの取得に失敗しました');
            }

            const data = await response.json();
            return data || [];

        } catch (error) {
            console.error('スプレッドシート取得エラー:', error);
            throw error;
        }
    }

    // IDマスターリストをレンダリング（階層対応）
    renderIdMasterList(folders, spreadsheets, currentFolderId = 'root') {
        const list = document.getElementById('id-master-list');
        
        let html = '';
        
        // 操作説明を追加
        html += '<div class="folder-instructions">';
        html += '<p><strong>💡 操作方法:</strong></p>';
        html += '<ul>';
        html += '<li>📁 <strong>フォルダ名をクリック</strong> → サブフォルダを表示</li>';
        html += '<li>📊 <strong>スプレッドシート/Excelファイル名をクリック</strong> → IDマスターとして選択</li>';
        html += '</ul>';
        html += '</div>';
        
        // パンくずリストを表示
        if (currentFolderId !== 'root' || this.breadcrumbs.length > 0) {
            html += '<div class="folder-breadcrumbs">';
            html += '<button class="breadcrumb-btn" onclick="mappingManager.showIdMasterModal(\'root\')">🏠 ルート</button>';
            
            this.breadcrumbs.forEach(breadcrumb => {
                html += ` <span class="breadcrumb-separator">></span> `;
                html += `<button class="breadcrumb-btn" onclick="mappingManager.showIdMasterModal('${breadcrumb.id}')">${this.escapeHtml(breadcrumb.name)}</button>`;
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
                <div class="list-item back-button" onclick="mappingManager.showIdMasterModal('${parentId}')">
                    <div class="item-icon">⬅️</div>
                    <div class="item-info">
                        <div class="item-name">戻る</div>
                        <div style="font-size: 0.8rem; color: #666;">上の階層に戻る</div>
                    </div>
                </div>
            `;
        }
        
        // フォルダを表示
        if (folders.length > 0) {
            html += '<div class="section-header" style="margin-top: 10px; font-weight: bold; color: #4a5568;">📁 フォルダ</div>';
            folders.forEach(folder => {
                html += `
                    <div class="list-item" onclick="mappingManager.showIdMasterModal('${folder.id}')" title="クリックしてフォルダを開く">
                        <div class="item-icon">📁</div>
                        <div class="item-info">
                            <div class="item-name">${this.escapeHtml(folder.name)}</div>
                            <div class="item-modified">フォルダを開く</div>
                        </div>
                    </div>
                `;
            });
        }
        
        // スプレッドシート・Excelファイルを表示
        if (spreadsheets.length === 0) {
            if (folders.length === 0) {
                html += '<p class="no-items">このフォルダにはスプレッドシート・Excelファイルがありません。</p>';
            }
        } else {
            html += '<div class="section-header" style="margin-top: 10px; font-weight: bold; color: #4a5568;">📊 スプレッドシート・Excelファイル</div>';
            spreadsheets.forEach(sheet => {
                const icon = sheet.fileType === 'excel' ? '📗' : '📊';
                const fileTypeName = sheet.fileType === 'excel' ? 'Excel' : 'スプレッドシート';
                html += `
                    <div class="list-item" data-id="${sheet.id}" data-name="${this.escapeHtml(sheet.name)}" data-type="${sheet.fileType}">
                        <div class="item-icon">${icon}</div>
                        <div class="item-info">
                            <div class="item-name">${this.escapeHtml(sheet.name)}</div>
                            <div class="item-modified">${fileTypeName} - 最終更新: ${new Date(sheet.modifiedTime).toLocaleDateString('ja-JP')}</div>
                        </div>
                    </div>
                `;
            });
            
            // スプレッドシート・Excelファイルのクリックイベントを設定
            setTimeout(() => {
                document.querySelectorAll('#id-master-list .list-item[data-id]').forEach(item => {
                    item.addEventListener('click', () => {
                        const fileId = item.getAttribute('data-id');
                        const fileName = item.getAttribute('data-name');
                        const fileType = item.getAttribute('data-type');
                        this.handleIdMasterSelection(fileId, fileName, fileType);
                        this.closeIdMasterModal();
                    });
                });
            }, 0);
        }
        
        list.innerHTML = html;
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // IDマスター選択モーダルを閉じる
    closeIdMasterModal() {
        const modal = document.getElementById('id-master-modal');
        if (modal) {
            modal.style.display = 'none';
        }
    }

    // IDマスタースプレッドシート/Excelファイルの処理
    async handleIdMasterSelection(fileId, fileName, fileType = 'spreadsheet') {
        try {
            showProgress('IDマスターを読み込み中...', 0);
            
            this.idMasterSpreadsheetId = fileId;
            this.idMasterSpreadsheetName = fileName;
            
            let materialData, processData, implementerData;
            
            if (fileType === 'excel') {
                // Excelファイルの場合
                console.log('📗 Excelファイルとして処理します');
                const excelData = await this.fetchExcelData(fileId);
                
                // Excelファイルを解析
                const workbook = new ExcelJS.Workbook();
                const arrayBuffer = this.base64ToArrayBuffer(excelData.content);
                await workbook.xlsx.load(arrayBuffer);
                
                // 各シートを取得
                const materialSheet = workbook.getWorksheet('素材IDマスター');
                if (!materialSheet) {
                    throw new Error('「素材IDマスター」という名前のシートが見つかりません');
                }
                
                const processSheet = workbook.getWorksheet('加工IDマスター');
                if (!processSheet) {
                    throw new Error('「加工IDマスター」という名前のシートが見つかりません');
                }
                
                const implementerSheet = workbook.getWorksheet('実施者IDマスター');
                if (!implementerSheet) {
                    throw new Error('「実施者IDマスター」という名前のシートが見つかりません');
                }
                
                // ExcelJSシートから配列データに変換
                materialData = this.excelSheetToArray(materialSheet);
                processData = this.excelSheetToArray(processSheet);
                implementerData = this.excelSheetToArray(implementerSheet);
                
            } else {
                // Googleスプレッドシートの場合
                console.log('📊 Googleスプレッドシートとして処理します');
                materialData = await this.fetchSheetData(fileId, '素材IDマスター');
                processData = await this.fetchSheetData(fileId, '加工IDマスター');
                implementerData = await this.fetchSheetData(fileId, '実施者IDマスター');
            }
            
            // データを解析
            this.materialMapping = await this.parseMaterialData(materialData);
            console.log('✅ 素材IDマスター読み込み完了:', Object.keys(this.materialMapping).length, '件');

            this.processMapping = await this.parseProcessData(processData);
            console.log('✅ 加工IDマスター読み込み完了:', Object.keys(this.processMapping).length, '件');

            this.implementerMapping = await this.parseImplementerData(implementerData);
            console.log('✅ 実施者IDマスター読み込み完了:', Object.keys(this.implementerMapping).length, '件');

            // UI更新
            const fileTypeName = fileType === 'excel' ? 'Excelファイル' : 'スプレッドシート';
            const infoBox = document.getElementById('id-master-info');
            infoBox.innerHTML = `
                <div class="mapping-file-success">
                    <h4>✅ IDマスターが読み込まれました</h4>
                    <p><strong>${fileTypeName}名:</strong> ${fileName}</p>
                    <div class="mapping-section">
                        <p><strong>📦 素材IDマスター:</strong> ${Object.keys(this.materialMapping).length} 件</p>
                        <div class="mapping-preview">
                            ${this.generateMaterialMappingPreview(this.materialMapping, 3)}
                        </div>
                    </div>
                    <div class="mapping-section">
                        <p><strong>⚙️ 加工IDマスター:</strong> ${Object.keys(this.processMapping).length} 件</p>
                        <div class="mapping-preview">
                            ${this.generateMappingPreview(this.processMapping, 3)}
                        </div>
                    </div>
                    <div class="mapping-section">
                        <p><strong>👤 実施者IDマスター:</strong> ${Object.keys(this.implementerMapping).length} 件</p>
                        <div class="mapping-preview">
                            ${this.generateMappingPreview(this.implementerMapping, 3)}
                        </div>
                    </div>
                </div>
            `;
            infoBox.classList.add('active');

            hideProgress();
            this.checkAllMappingsLoaded();

        } catch (error) {
            hideProgress();
            showError('IDマスターの読み込みに失敗しました: ' + error.message);
            console.error('IDマスターエラー:', error);
        }
    }

    // Excelファイルをダウンロード
    async fetchExcelData(fileId) {
        try {
            const response = await fetch(
                `${CONFIG.API_BASE_URL}/api/files/${fileId}/excel-content`,
                {
                    method: 'GET',
                    credentials: 'include'
                }
            );

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'Excelファイルのダウンロードに失敗しました');
            }

            return await response.json();

        } catch (error) {
            throw new Error(`Excelファイルの取得に失敗: ${error.message}`);
        }
    }

    // Base64をArrayBufferに変換
    base64ToArrayBuffer(base64) {
        const binaryString = window.atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes.buffer;
    }

    // ExcelJSシートを配列データに変換
    excelSheetToArray(worksheet) {
        const jsonData = [];
        const rowCount = worksheet.rowCount;
        const colCount = worksheet.columnCount;
        
        console.log(`📊 ワークシート情報: ${rowCount}行 x ${colCount}列`);
        
        for (let rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
            const row = worksheet.getRow(rowIndex);
            const rowData = [];
            
            for (let colIndex = 1; colIndex <= colCount; colIndex++) {
                const cell = row.getCell(colIndex);
                const cellValue = this.normalizeExcelValue(cell.value);
                rowData[colIndex - 1] = cellValue;
            }
            
            jsonData.push(rowData);
        }
        
        return jsonData;
    }

    // ExcelJSセル値の正規化
    normalizeExcelValue(value) {
        if (value === null || value === undefined) return null;
        if (typeof value !== 'object') return value;
        
        // ExcelJSの特殊な値タイプを処理
        if (value.richText) {
            return value.richText.map(part => part.text).join('');
        } else if (value.text) {
            return value.text;
        } else if (value.result !== undefined) {
            return value.result;
        }
        
        return value;
    }

    // Google Sheets APIでシートデータを取得
    async fetchSheetData(spreadsheetId, sheetName) {
        try {
            const response = await fetch(
                `${CONFIG.API_BASE_URL}/api/spreadsheets/${spreadsheetId}/sheets/${encodeURIComponent(sheetName)}/data`,
                {
                    method: 'GET',
                    credentials: 'include'
                }
            );

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `シート「${sheetName}」の読み込みに失敗しました`);
            }

            const data = await response.json();
            return data.values || [];

        } catch (error) {
            throw new Error(`シート「${sheetName}」の取得に失敗: ${error.message}`);
        }
    }

    // 素材IDマスターデータを解析
    async parseMaterialData(values) {
        try {
            if (!values || values.length < 2) {
                throw new Error('素材IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = values[0];
            const idColumnIndex = headerRow.findIndex(header => 
                header && header.toString().trim() === '素材ID'
            );
            const nameColumnIndex = headerRow.findIndex(header => 
                header && header.toString().trim() === '素材名'
            );
            const categoryColumnIndex = headerRow.findIndex(header => 
                header && header.toString().trim() === '素材区分'
            );

            if (idColumnIndex === -1) {
                throw new Error('素材IDマスターに「素材ID」列が見つかりません');
            }
            if (nameColumnIndex === -1) {
                throw new Error('素材IDマスターに「素材名」列が見つかりません');
            }
            if (categoryColumnIndex === -1) {
                throw new Error('素材IDマスターに「素材区分」列が見つかりません');
            }

            // データ行を処理してマッピングオブジェクトを作成
            const mapping = {};
            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const id = row[idColumnIndex];
                const name = row[nameColumnIndex];
                const category = row[categoryColumnIndex];
                
                if (id && name && category) {
                    mapping[id.toString().trim()] = {
                        name: name.toString().trim(),
                        category: category.toString().trim()
                    };
                }
            }

            if (Object.keys(mapping).length === 0) {
                throw new Error('素材IDマスターに有効なデータが見つかりませんでした');
            }

            return mapping;

        } catch (error) {
            throw new Error('素材IDマスターの解析に失敗しました: ' + error.message);
        }
    }

    // 加工IDマスターデータを解析
    async parseProcessData(values) {
        try {
            if (!values || values.length < 2) {
                throw new Error('加工IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = values[0];
            console.log('🔍 加工IDマスターヘッダー行:', headerRow);
            
            const normalizeHeaderName = (name) => {
                if (!name) return '';
                return name.toString().trim()
                    .replace(/\s+/g, '')
                    .toLowerCase()
                    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
            };
            
            const idColumnIndex = headerRow.findIndex(header => 
                normalizeHeaderName(header) === normalizeHeaderName('加工ID')
            );
            const nameColumnIndex = headerRow.findIndex(header => {
                const normalizedHeader = normalizeHeaderName(header);
                return normalizedHeader === normalizeHeaderName('加工方法名') ||
                       normalizedHeader === normalizeHeaderName('加工方法') ||
                       normalizedHeader.includes('加工方法');
            });

            if (idColumnIndex === -1) {
                throw new Error('加工IDマスターに「加工ID」列が見つかりません');
            }
            if (nameColumnIndex === -1) {
                throw new Error('加工IDマスターに「加工方法名」列が見つかりません');
            }

            // データ行を処理してマッピングオブジェクトを作成
            const mapping = {};
            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const id = row[idColumnIndex];
                const name = row[nameColumnIndex];
                
                if (id && name) {
                    mapping[id.toString().trim()] = name.toString().trim();
                }
            }

            if (Object.keys(mapping).length === 0) {
                throw new Error('加工IDマスターに有効なデータが見つかりませんでした');
            }

            return mapping;

        } catch (error) {
            throw new Error('加工IDマスターの解析に失敗しました: ' + error.message);
        }
    }

    // 実施者IDマスターデータを解析
    async parseImplementerData(values) {
        try {
            if (!values || values.length < 2) {
                throw new Error('実施者IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = values[0];
            console.log('🔍 実施者IDマスターヘッダー行:', headerRow);
            
            const normalizeHeaderName = (name) => {
                if (!name) return '';
                return name.toString().trim()
                    .replace(/\s+/g, '')
                    .toLowerCase()
                    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
            };
            
            const idColumnIndex = headerRow.findIndex(header => 
                normalizeHeaderName(header) === normalizeHeaderName('実施者ID')
            );
            const nameColumnIndex = headerRow.findIndex(header => 
                normalizeHeaderName(header) === normalizeHeaderName('実施者名')
            );

            if (idColumnIndex === -1) {
                throw new Error('実施者IDマスターに「実施者ID」列が見つかりません');
            }
            if (nameColumnIndex === -1) {
                throw new Error('実施者IDマスターに「実施者名」列が見つかりません');
            }

            // データ行を処理してマッピングオブジェクトを作成
            const mapping = {};
            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const id = row[idColumnIndex];
                const name = row[nameColumnIndex];
                
                if (id && name) {
                    mapping[id.toString().trim()] = name.toString().trim();
                }
            }

            if (Object.keys(mapping).length === 0) {
                throw new Error('実施者IDマスターに有効なデータが見つかりませんでした');
            }

            return mapping;

        } catch (error) {
            throw new Error('実施者IDマスターの解析に失敗しました: ' + error.message);
        }
    }

    // すべての対応表が読み込まれたかチェック
    checkAllMappingsLoaded() {
        if (this.materialMapping && this.processMapping && this.implementerMapping) {
            // 処理実行ボタンを有効化するためのチェック関数を呼び出し
            if (typeof checkProcessButtonState === 'function') {
                checkProcessButtonState();
            }
            
            // すべての対応表の読み込み完了メッセージを表示
            if (typeof showMessage === 'function') {
                showMessage('✅ IDマスターの読み込みが完了しました', 'success');
            } else {
                console.log('✅ IDマスターの読み込みが完了しました');
            }
        }
    }

    // 素材マッピングプレビューの生成
    generateMaterialMappingPreview(mapping, maxItems = 3) {
        const entries = Object.entries(mapping).slice(0, maxItems);
        const preview = entries.map(([id, data]) => `${id} → ${data.name}(${data.category})`).join('<br>');
        const remaining = Object.keys(mapping).length - maxItems;
        
        return preview + (remaining > 0 ? `<br>...他 ${remaining} 件` : '');
    }

    // マッピングプレビューの生成
    generateMappingPreview(mapping, maxItems = 3) {
        const entries = Object.entries(mapping).slice(0, maxItems);
        const preview = entries.map(([id, name]) => `${id} → ${name}`).join('<br>');
        const remaining = Object.keys(mapping).length - maxItems;
        
        return preview + (remaining > 0 ? `<br>...他 ${remaining} 件` : '');
    }

    // IDから素材データ（名前と区分）への変換
    getMaterialData(materialId) {
        if (!this.materialMapping) {
            return { name: '該当なし', category: '該当なし' };
        }
        const materialData = this.materialMapping[materialId];
        if (!materialData) {
            return { name: '該当なし', category: '該当なし' };
        }
        return materialData;
    }

    // 従来の互換性のためのメソッド（非推奨）
    getMaterialName(materialId) {
        const materialData = this.getMaterialData(materialId);
        return materialData.name;
    }

    // IDから加工方法名への変換
    getProcessName(processId) {
        if (!this.processMapping) {
            return '該当なし';
        }
        return this.processMapping[processId] || '該当なし';
    }

    // IDから実施者名への変換
    getImplementerName(implementerId) {
        if (!this.implementerMapping) {
            return '該当なし';
        }
        return this.implementerMapping[implementerId] || '該当なし';
    }

    // 対応表が準備完了かチェック
    isReady() {
        return !!(this.materialMapping && this.processMapping && this.implementerMapping);
    }

    // リセット
    reset() {
        this.materialMapping = null;
        this.processMapping = null;
        this.implementerMapping = null;
        this.idMasterSpreadsheetId = null;
        this.idMasterSpreadsheetName = null;

        // UI リセット
        const idMasterInfo = document.getElementById('id-master-info');
        
        if (idMasterInfo) idMasterInfo.classList.remove('active');
    }
}

// モーダルを閉じるグローバル関数
function closeIdMasterModal() {
    if (mappingManager) {
        mappingManager.closeIdMasterModal();
    }
}

// グローバルインスタンス
const mappingManager = new MappingTableManager();