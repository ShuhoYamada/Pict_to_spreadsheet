// ID対応表管理クラス - ExcelJS対応版
class MappingTableManager {
    constructor() {
        this.materialMapping = null;
        this.processMapping = null;
        this.materialMappingFile = null;
        this.processMappingFile = null;
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

        // 素材ID対応表選択ボタン
        const materialButton = document.getElementById('select-material-mapping-button');
        if (materialButton) {
            materialButton.addEventListener('click', () => {
                document.getElementById('material-mapping-file').click();
            });
        }

        // 加工ID対応表選択ボタン
        const processButton = document.getElementById('select-process-mapping-button');
        if (processButton) {
            processButton.addEventListener('click', () => {
                document.getElementById('process-mapping-file').click();
            });
        }

        // ファイル選択イベント
        const materialFileInput = document.getElementById('material-mapping-file');
        if (materialFileInput) {
            materialFileInput.addEventListener('change', (event) => {
                this.handleMaterialMappingFile(event.target.files[0]);
            });
        }

        const processFileInput = document.getElementById('process-mapping-file');
        if (processFileInput) {
            processFileInput.addEventListener('change', (event) => {
                this.handleProcessMappingFile(event.target.files[0]);
            });
        }

        this.eventListenersSetup = true;
        console.log('✅ マッピングマネージャーのイベントリスナー設定完了');
    }

    // 素材ID対応表ファイルの処理
    async handleMaterialMappingFile(file) {
        if (!file) return;

        try {
            showProgress('素材ID対応表を読み込み中...', 0);
            
            this.materialMappingFile = file;
            // 素材名と素材区分の両方を取得
            const mappingData = await this.parseMaterialExcelFile(file);
            this.materialMapping = mappingData;

            // UI更新
            const infoBox = document.getElementById('material-mapping-info');
            infoBox.innerHTML = `
                <div class="mapping-file-success">
                    <h4>✅ 素材ID対応表が読み込まれました</h4>
                    <p><strong>ファイル名:</strong> ${file.name}</p>
                    <p><strong>データ数:</strong> ${Object.keys(mappingData).length} 件</p>
                    <div class="mapping-preview">
                        <strong>プレビュー:</strong>
                        ${this.generateMaterialMappingPreview(mappingData, 3)}
                    </div>
                </div>
            `;
            infoBox.classList.add('active');

            hideProgress();
            this.checkAllMappingsLoaded();

        } catch (error) {
            hideProgress();
            showError('素材ID対応表の読み込みに失敗しました: ' + error.message);
            console.error('素材ID対応表エラー:', error);
        }
    }

    // 加工ID対応表ファイルの処理
    async handleProcessMappingFile(file) {
        if (!file) return;

        try {
            showProgress('加工ID対応表を読み込み中...', 0);
            
            this.processMappingFile = file;
            const mappingData = await this.parseExcelFile(file, '加工ID', '加工方法名');
            this.processMapping = mappingData;

            // UI更新
            const infoBox = document.getElementById('process-mapping-info');
            infoBox.innerHTML = `
                <div class="mapping-file-success">
                    <h4>✅ 加工ID対応表が読み込まれました</h4>
                    <p><strong>ファイル名:</strong> ${file.name}</p>
                    <p><strong>データ数:</strong> ${Object.keys(mappingData).length} 件</p>
                    <div class="mapping-preview">
                        <strong>プレビュー:</strong>
                        ${this.generateMappingPreview(mappingData, 3)}
                    </div>
                </div>
            `;
            infoBox.classList.add('active');

            hideProgress();
            this.checkAllMappingsLoaded();

        } catch (error) {
            hideProgress();
            showError('加工ID対応表の読み込みに失敗しました: ' + error.message);
            console.error('加工ID対応表エラー:', error);
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

    // Excelファイルの解析（ExcelJS使用）
    async parseExcelFile(file, idColumnName, nameColumnName) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            
            // ExcelJSライブラリを使用してExcelファイルを解析
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            // 最初のシートを取得
            const worksheet = workbook.worksheets[0];
            
            // データ範囲を取得してJSONに変換
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
            
            console.log('📊 解析されたデータ行数:', jsonData.length);
            
            if (jsonData.length < 2) {
                throw new Error('ファイルにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
            }

            // デバッグ: ヘッダー行の内容を確認
            console.log('🔍 デバッグ - ヘッダー行:', jsonData[0]);
            console.log('🔍 デバッグ - 探しているID列名:', idColumnName);
            console.log('🔍 デバッグ - 探している名前列名:', nameColumnName);
            
            // 各ヘッダー値を詳細に確認
            const headerRow = jsonData[0];
            headerRow.forEach((header, index) => {
                console.log(`🔍 列${index}:`, typeof header, '|', header, '|', 
                           header ? header.toString().trim() : '(empty)');
            });
            
            // より柔軟なヘッダー検索（大文字小文字、スペース、全角半角を無視）
            const normalizeHeaderName = (name) => {
                if (!name) return '';
                return name.toString().trim()
                    .replace(/\s+/g, '')  // 全てのスペースを削除
                    .toLowerCase()        // 小文字に変換
                    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)); // 全角→半角
            };
            
            const targetIdName = normalizeHeaderName(idColumnName);
            const targetNameName = normalizeHeaderName(nameColumnName);
            
            console.log('🔍 正規化された検索対象:', {
                id: targetIdName,
                name: targetNameName
            });
            
            // ヘッダー行から列インデックスを特定
            const idColumnIndex = headerRow.findIndex(header => 
                normalizeHeaderName(header) === targetIdName
            );
            const nameColumnIndex = headerRow.findIndex(header => {
                const normalizedHeader = normalizeHeaderName(header);
                const normalizedTarget = normalizeHeaderName(nameColumnName);
                
                // 完全一致
                if (normalizedHeader === normalizedTarget) return true;
                
                // 加工方法関連の特別処理
                if (normalizedTarget.includes('加工') && normalizedTarget.includes('方法')) {
                    return normalizedHeader === '加工方法' || 
                           normalizedHeader === '加工方法名' ||
                           normalizedHeader.includes('加工方法');
                }
                
                return false;
            });

            if (idColumnIndex === -1) {
                throw new Error(`「${idColumnName}」列が見つかりません`);
            }
            if (nameColumnIndex === -1) {
                throw new Error(`「${nameColumnName}」列が見つかりません`);
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
                throw new Error('有効なデータが見つかりませんでした');
            }

            return mapping;

        } catch (error) {
            throw new Error('ファイルの解析に失敗しました: ' + error.message);
        }
    }

    // 素材Excelファイルの解析（素材名と素材区分の両方を取得）
    async parseMaterialExcelFile(file) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            
            // ExcelJSライブラリを使用してExcelファイルを解析
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            
            // 最初のシートを取得
            const worksheet = workbook.worksheets[0];
            
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
                throw new Error('ファイルにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
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
                throw new Error('「素材ID」列が見つかりません');
            }
            if (nameColumnIndex === -1) {
                throw new Error('「素材名」列が見つかりません');
            }
            if (categoryColumnIndex === -1) {
                throw new Error('「素材区分」列が見つかりません');
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
                throw new Error('有効なデータが見つかりませんでした');
            }

            return mapping;

        } catch (error) {
            throw new Error('ファイルの解析に失敗しました: ' + error.message);
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

    // すべての対応表が読み込まれたかチェック
    checkAllMappingsLoaded() {
        if (this.materialMapping && this.processMapping) {
            // 処理実行ボタンを有効化するためのチェック関数を呼び出し
            if (typeof checkProcessButtonState === 'function') {
                checkProcessButtonState();
            }
            
            // すべての対応表の読み込み完了メッセージを表示
            if (typeof showMessage === 'function') {
                showMessage('✅ すべての対応表の読み込みが完了しました', 'success');
            } else {
                console.log('✅ すべての対応表の読み込みが完了しました');
            }
        }
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

    // 対応表が準備完了かチェック
    isReady() {
        return !!(this.materialMapping && this.processMapping);
    }

    // リセット
    reset() {
        this.materialMapping = null;
        this.processMapping = null;
        this.materialMappingFile = null;
        this.processMappingFile = null;

        // UI リセット
        const materialInfo = document.getElementById('material-mapping-info');
        const processInfo = document.getElementById('process-mapping-info');
        const materialFile = document.getElementById('material-mapping-file');
        const processFile = document.getElementById('process-mapping-file');
        
        if (materialInfo) materialInfo.classList.remove('active');
        if (processInfo) processInfo.classList.remove('active');
        if (materialFile) materialFile.value = '';
        if (processFile) processFile.value = '';
    }
}

// グローバルインスタンス
const mappingManager = new MappingTableManager();