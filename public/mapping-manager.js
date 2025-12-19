// ID対応表管理クラス - ExcelJS対応版
class MappingTableManager {
    constructor() {
        this.materialMapping = null;
        this.processMapping = null;
        this.implementerMapping = null;
        this.idMasterFile = null;
        this.eventListenersSetup = false; // 重複防止フラグ
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
            idMasterButton.addEventListener('click', () => {
                document.getElementById('id-master-file').click();
            });
        }

        // ファイル選択イベント
        const idMasterFileInput = document.getElementById('id-master-file');
        if (idMasterFileInput) {
            idMasterFileInput.addEventListener('change', (event) => {
                this.handleIdMasterFile(event.target.files[0]);
            });
        }

        this.eventListenersSetup = true;
        console.log('✅ マッピングマネージャーのイベントリスナー設定完了');
    }

    // IDマスターファイルの処理
    async handleIdMasterFile(file) {
        if (!file) return;

        try {
            showProgress('IDマスターを読み込み中...', 0);
            
            this.idMasterFile = file;
            
            // Excelファイルを読み込んで2つのシートから情報を取得
            const arrayBuffer = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);

            // 素材IDマスターシートを取得
            const materialSheet = workbook.getWorksheet('素材IDマスター');
            if (!materialSheet) {
                throw new Error('「素材IDマスター」という名前のシートが見つかりません');
            }

            // 加工IDマスターシートを取得
            const processSheet = workbook.getWorksheet('加工IDマスター');
            if (!processSheet) {
                throw new Error('「加工IDマスター」という名前のシートが見つかりません');
            }

            // 実施者IDマスターシートを取得
            const implementerSheet = workbook.getWorksheet('実施者IDマスター');
            if (!implementerSheet) {
                throw new Error('「実施者IDマスター」という名前のシートが見つかりません');
            }

            // 素材データを解析
            this.materialMapping = await this.parseMaterialSheet(materialSheet);
            console.log('✅ 素材IDマスター読み込み完了:', Object.keys(this.materialMapping).length, '件');

            // 加工データを解析
            this.processMapping = await this.parseProcessSheet(processSheet);
            console.log('✅ 加工IDマスター読み込み完了:', Object.keys(this.processMapping).length, '件');

            // 実施者データを解析
            this.implementerMapping = await this.parseImplementerSheet(implementerSheet);
            console.log('✅ 実施者IDマスター読み込み完了:', Object.keys(this.implementerMapping).length, '件');

            // UI更新
            const infoBox = document.getElementById('id-master-info');
            infoBox.innerHTML = `
                <div class="mapping-file-success">
                    <h4>✅ IDマスターが読み込まれました</h4>
                    <p><strong>ファイル名:</strong> ${file.name}</p>
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

    // ExcelJSセル値の正規化ヘルパー
    normalizeExcelValue(value) {
        if (value === null || value === undefined) return null;
        if (typeof value !== 'object') return value;
        
        // ExcelJSの特殊な値タイプを処理
        if (value.richText) {
            // リッチテキストの場合、テキスト部分を抽出
            return value.richText.map(part => part.text).join('');
        } else if (value.text) {
            // テキストオブジェクトの場合
            return value.text;
        } else if (value.result !== undefined) {
            // 数式の結果値
            return value.result;
        }
        
        return value;
    }

    // 素材IDマスターシートを解析
    async parseMaterialSheet(worksheet) {
        try {
            // データ範囲を取得してJSONに変換
            const jsonData = [];
            const rowCount = worksheet.rowCount;
            const colCount = worksheet.columnCount;
            
            console.log(`📊 素材ワークシート情報: ${rowCount}行 x ${colCount}列`);
            
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
            
            if (jsonData.length < 2) {
                throw new Error('素材IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = jsonData[0];
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
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
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

    // 加工IDマスターシートを解析
    async parseProcessSheet(worksheet) {
        try {
            // データ範囲を取得してJSONに変換
            const jsonData = [];
            const rowCount = worksheet.rowCount;
            const colCount = worksheet.columnCount;
            
            console.log(`📊 加工ワークシート情報: ${rowCount}行 x ${colCount}列`);
            
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
            
            if (jsonData.length < 2) {
                throw new Error('加工IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = jsonData[0];
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
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
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

    // 実施者IDマスターシートを解析
    async parseImplementerSheet(worksheet) {
        try {
            // データ範囲を取得してJSONに変換
            const jsonData = [];
            const rowCount = worksheet.rowCount;
            const colCount = worksheet.columnCount;
            
            console.log(`📊 実施者ワークシート情報: ${rowCount}行 x ${colCount}列`);
            
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
            
            if (jsonData.length < 2) {
                throw new Error('実施者IDマスターにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // ヘッダー行から列インデックスを特定
            const headerRow = jsonData[0];
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
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
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
        this.idMasterFile = null;

        // UI リセット
        const idMasterInfo = document.getElementById('id-master-info');
        const idMasterFile = document.getElementById('id-master-file');
        
        if (idMasterInfo) idMasterInfo.classList.remove('active');
        if (idMasterFile) idMasterFile.value = '';
    }
}

// グローバルインスタンス
const mappingManager = new MappingTableManager();