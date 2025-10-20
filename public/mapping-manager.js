// ID対応表管理クラス
class MappingTableManager {
    constructor() {
        this.materialMapping = null;
        this.processMapping = null;
        this.materialMappingFile = null;
        this.processMappingFile = null;
    }

    // 初期化
    initialize() {
        this.setupEventListeners();
    }

    // イベントリスナーの設定
    setupEventListeners() {
        // 素材ID対応表選択ボタン
        document.getElementById('select-material-mapping-button').addEventListener('click', () => {
            document.getElementById('material-mapping-file').click();
        });

        // 加工ID対応表選択ボタン
        document.getElementById('select-process-mapping-button').addEventListener('click', () => {
            document.getElementById('process-mapping-file').click();
        });

        // ファイル選択イベント
        document.getElementById('material-mapping-file').addEventListener('change', (event) => {
            this.handleMaterialMappingFile(event.target.files[0]);
        });

        document.getElementById('process-mapping-file').addEventListener('change', (event) => {
            this.handleProcessMappingFile(event.target.files[0]);
        });
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

    // Excelファイルの解析
    async parseExcelFile(file, idColumnName, nameColumnName) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    // SheetJSライブラリを使用してExcelファイルを解析
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 最初のシートを取得
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // JSONに変換
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    if (jsonData.length < 2) {
                        throw new Error('ファイルにデータが不足しています（ヘッダー + 最低1行のデータが必要）');
                    }

                    // ヘッダー行から列インデックスを特定
                    const headerRow = jsonData[0];
                    const idColumnIndex = headerRow.findIndex(header => 
                        header && header.toString().trim() === idColumnName
                    );
                    const nameColumnIndex = headerRow.findIndex(header => 
                        header && header.toString().trim() === nameColumnName
                    );

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

                    resolve(mapping);

                } catch (error) {
                    reject(new Error('ファイルの解析に失敗しました: ' + error.message));
                }
            };

            reader.onerror = () => {
                reject(new Error('ファイルの読み込みに失敗しました'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    // 素材Excelファイルの解析（素材名と素材区分の両方を取得）
    async parseMaterialExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    // SheetJSライブラリを使用してExcelファイルを解析
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 最初のシートを取得
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // JSONに変換
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
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

                    resolve(mapping);

                } catch (error) {
                    reject(new Error('ファイルの解析に失敗しました: ' + error.message));
                }
            };

            reader.onerror = () => {
                reject(new Error('ファイルの読み込みに失敗しました'));
            };

            reader.readAsArrayBuffer(file);
        });
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
        document.getElementById('material-mapping-info').classList.remove('active');
        document.getElementById('process-mapping-info').classList.remove('active');
        document.getElementById('material-mapping-file').value = '';
        document.getElementById('process-mapping-file').value = '';
    }
}

// グローバルインスタンス
const mappingManager = new MappingTableManager();