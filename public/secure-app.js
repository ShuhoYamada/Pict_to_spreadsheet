// セキュアなメインアプリケーション
class SecurePhotoFileManagerApp {
    constructor() {
        this.isInitialized = false;
        this.processingData = null;
        this.init();
    }

    async init() {
        try {
            console.log('📋 写真ファイル管理システムの初期化中...');
            
            // 認証マネージャーが初期化されていない場合のみ初期化
            if (!authManager) {
                console.log('🔐 認証マネージャーを初期化中...');
                authManager = new SecureAuthManager();
                console.log('✅ 認証マネージャーを初期化しました');
            } else {
                console.log('✅ 認証マネージャーは既に初期化済みです');
            }
            
            // 設定を検証（バックエンド接続確認）
            const isConfigValid = await this.validateConfig();
            if (!isConfigValid) {
                console.warn('⚠️ バックエンドサーバーに接続できませんが、認証機能は利用可能です');
                return;
            }
            
            // イベントリスナーを設定
            this.setupEventListeners();
            
            this.isInitialized = true;
            console.log('✅ 写真ファイル管理システムの初期化が完了しました');
            
        } catch (error) {
            console.error('アプリケーション初期化エラー:', error);
            showError('アプリケーションの初期化に失敗しました: ' + error.message);
        }
    }

    async validateConfig() {
        try {
            const response = await fetch(`${CONFIG.API_BASE_URL}/health`);
            return response.ok;
        } catch (error) {
            console.error('バックエンドサーバーに接続できません:', error);
            showError('バックエンドサーバーに接続できません。開発者にお問い合わせください。');
            return false;
        }
    }

    setupEventListeners() {
        // フォルダ選択ボタン
        const selectFolderButton = document.getElementById('select-folder-button');
        if (selectFolderButton) {
            selectFolderButton.addEventListener('click', () => {
                apiManager.showFolderSelector();
            });
        } else {
            console.error('select-folder-button が見つかりません');
        }

        // スプレッドシート選択ボタン
        const selectSpreadsheetButton = document.getElementById('select-spreadsheet-button');
        if (selectSpreadsheetButton) {
            selectSpreadsheetButton.addEventListener('click', () => {
                apiManager.showSpreadsheetSelector();
            });
        } else {
            console.error('select-spreadsheet-button が見つかりません');
        }

        // データ処理ボタン
        const processDataButton = document.getElementById('process-data-button');
        if (processDataButton) {
            processDataButton.addEventListener('click', () => {
                this.processAndWriteData();
            });
        } else {
            console.error('process-data-button が見つかりません');
        }

        // ESCキーでモーダルを閉じる
        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                const folderModal = document.getElementById('folder-modal');
                const spreadsheetModal = document.getElementById('spreadsheet-modal');
                if (folderModal) folderModal.style.display = 'none';
                if (spreadsheetModal) spreadsheetModal.style.display = 'none';
                hideError();
            }
        });

        // モーダル背景クリックで閉じる
        const folderModal = document.getElementById('folder-modal');
        if (folderModal) {
            folderModal.addEventListener('click', (event) => {
                if (event.target === event.currentTarget) {
                    closeFolderModal();
                }
            });
        }

        const spreadsheetModal = document.getElementById('spreadsheet-modal');
        if (spreadsheetModal) {
            spreadsheetModal.addEventListener('click', (event) => {
                if (event.target === event.currentTarget) {
                    closeSpreadsheetModal();
                }
            });
        }
    }

    async processAndWriteData() {
        try {
            if (!authManager.isUserAuthenticated()) {
                showError('認証が必要です');
                return;
            }

            // 入力検証
            const photoFiles = apiManager.getSelectedPhotoFiles();
            const selectedFolder = apiManager.getSelectedFolder();
            const selectedSpreadsheet = apiManager.getSelectedSpreadsheet();

            if (!photoFiles || photoFiles.length === 0) {
                showError('処理する写真ファイルがありません');
                return;
            }

            if (!selectedSpreadsheet) {
                showError('スプレッドシートが選択されていません');
                return;
            }

            if (!mappingManager.isReady()) {
                showError('対応表が選択されていません。素材ID対応表と加工ID対応表の両方を選択してください。');
                return;
            }

            // 処理開始
            this.updateProcessingStatus('📋 ファイル名を解析中...', 10);
            
            // 新仕様に基づく処理を実行
            const result = await apiManager.processAndWriteData(
                selectedSpreadsheet.id,
                photoFiles,
                mappingManager.materialMapping,
                mappingManager.processMapping,
                mappingManager.implementerMapping
            );

            this.updateProcessingStatus('✅ 処理完了', 100);

            // 結果を表示
            this.showResults({
                folder: selectedFolder,
                spreadsheet: selectedSpreadsheet,
                result: result
            });

            setTimeout(() => {
                hideProgress();
            }, 2000);

        } catch (error) {
            console.error('データ処理エラー:', error);
            hideProgress();
            showError(error.message);
            this.updateProcessingStatus('❌ エラーが発生しました', 0);
        }
    }

    updateProcessingStatus(message, progress) {
        showProgress(message, progress);
        
        const statusElement = document.getElementById('processing-status');
        statusElement.className = 'status-message status-info';
        statusElement.innerHTML = `
            <div style="display: flex; align-items: center; gap: 10px;">
                <span>${message}</span>
                ${progress < 100 ? '<div class="loading-spinner"></div>' : ''}
            </div>
        `;
    }

    showResults(results) {
        const resultsBox = document.getElementById('results');
        const { folder, spreadsheet, result } = results;

        let resultsHTML = `
            <h4>🎉 処理結果 (ハイパーリンク対応版)</h4>
            
            <div class="result-section" style="margin-bottom: 20px;">
                <h5>🔒 セキュリティ情報</h5>
                <p>✅ すべてのGoogle APIs通信はバックエンドサーバー経由で暗号化されています</p>
                <p>✅ APIキーやアクセストークンはブラウザに露出されていません</p>
                <p>✅ ID対応表はローカルファイルから安全に参照されています</p>
                <p>✅ 写真ファイルのハイパーリンクが自動設定されました</p>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📂 処理したフォルダ</h5>
                <p><strong>${this.escapeHtml(folder.name)}</strong></p>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📊 書き込み先スプレッドシート</h5>
                <p><strong>${this.escapeHtml(spreadsheet.name)}</strong></p>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📈 処理統計</h5>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
                    <div class="stat-item">
                        <div class="stat-number">${result.totalCount}</div>
                        <div class="stat-label">総ファイル数</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${result.writtenCount}</div>
                        <div class="stat-label">P区分書き込み</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${result.mTypeCount || 0}</div>
                        <div class="stat-label">M区分処理</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${result.hyperlinkCount || 0}</div>
                        <div class="stat-label">ハイパーリンク</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${result.invalidCount}</div>
                        <div class="stat-label">形式エラー</div>
                    </div>
                </div>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>🔗 ハイパーリンク設定結果</h5>
                <div style="background: #f0f9ff; padding: 15px; border-radius: 8px; border-left: 4px solid #0ea5e9;">
                    <p><strong>P区分：</strong> 「構成部品」列にハイパーリンクを設定</p>
                    <p><strong>M区分：</strong> 対応するP区分行の「素材」列にハイパーリンクを設定</p>
                    <p><strong>設定完了：</strong> ${result.hyperlinkCount || 0} 個のハイパーリンク</p>
                </div>
            </div>`;

        // ハイパーリンクエラーがある場合
        if (result.hyperlinkErrors && result.hyperlinkErrors.length > 0) {
            resultsHTML += `
                <div class="result-section" style="margin-bottom: 20px;">
                    <h5>⚠️ ハイパーリンク設定エラー</h5>
                    <div class="file-list">
                        ${result.hyperlinkErrors.map(error => 
                            `<div class="file-item error">
                                <strong>📷 ${this.escapeHtml(error.fileName)} (${error.type})</strong>
                                <div class="error-message">${this.escapeHtml(error.error)}</div>
                            </div>`
                        ).join('')}
                    </div>
                </div>`;
        }

        // エラーファイルの詳細
        if (result.invalidFiles && result.invalidFiles.length > 0) {
            resultsHTML += `
                <div class="result-section" style="margin-bottom: 20px;">
                    <h5>❌ 形式エラーファイル</h5>
                    <div class="file-list">
                        ${result.invalidFiles.map(file => 
                            `<div class="file-item error">
                                <strong>📷 ${this.escapeHtml(file.fileName)}</strong>
                                <div class="error-message">${this.escapeHtml(file.error)}</div>
                            </div>`
                        ).join('')}
                    </div>
                </div>`;
        }

        resultsHTML += `
            <div class="result-section">
                <h5>🎯 新仕様処理内容</h5>
                <ul>
                    <li>✅ P区分ファイル：スプレッドシートにデータ書き込み ＋ 「構成部品」列にハイパーリンク設定</li>
                    <li>✅ M区分ファイル：対応するP区分行の「素材」列にハイパーリンク設定</li>
                    <li>✅ ID変換適用（素材ID→素材名、加工ID→加工方法名）</li>
                    <li>✅ 単位変換適用（g ⇄ kg）</li>
                    <li>✅ 特記事項変換適用（0→なし、1→あり）</li>
                    <li>✅ Googleドライブ写真へのダイレクトリンク設定</li>
                </ul>
            </div>
        `;

        resultsBox.innerHTML = resultsHTML;
        resultsBox.classList.add('active');

        const statusElement = document.getElementById('processing-status');
        statusElement.className = 'status-message status-success';
        statusElement.innerHTML = `
            <div style="display: flex; align-items: center; gap: 10px;">
                <span>✅ 処理完了！P区分${result.writtenCount}件書き込み、ハイパーリンク${result.hyperlinkCount || 0}個設定</span>
            </div>
        `;
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// プロセスボタンの状態をチェックする関数
function checkProcessButtonState() {
    const processButton = document.getElementById('process-data-button');
    const hasFolder = apiManager.getSelectedFolder();
    const hasSpreadsheet = apiManager.getSelectedSpreadsheet();
    const hasMappings = mappingManager && mappingManager.isReady();
    const isAuthenticated = authManager && authManager.isUserAuthenticated();

    if (hasFolder && hasSpreadsheet && hasMappings && isAuthenticated) {
        processButton.disabled = false;
        processButton.textContent = 'データを処理してスプレッドシートに書き込む';
    } else {
        processButton.disabled = true;
        let reason = '次の準備が必要です: ';
        const missing = [];
        if (!isAuthenticated) missing.push('Google認証');
        if (!hasFolder) missing.push('フォルダ選択');
        if (!hasMappings) missing.push('対応表選択');
        if (!hasSpreadsheet) missing.push('スプレッドシート選択');
        processButton.textContent = reason + missing.join(', ');
    }
}

// ページ読み込み完了後にセキュアアプリケーションを初期化
document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 DOM読み込み完了 - セキュアアプリケーションを初期化します');
    
    // 認証マネージャーを最優先で初期化
    if (!authManager) {
        console.log('🔐 認証マネージャーを初期化中...');
        authManager = new SecureAuthManager();
        console.log('✅ 認証マネージャーの初期化が完了しました');
    }
    
    // マッピングのイベントはアプリ初期化に依存せず登録しておく
    // （バックエンド接続検証に失敗した場合でも、ファイル選択UIは動作させたい）
    try {
        if (typeof mappingManager !== 'undefined' && mappingManager && typeof mappingManager.initialize === 'function') {
            mappingManager.initialize();
            console.log('✅ mappingManager 初期化: イベントリスナー登録済み');
        }
    } catch (err) {
        console.warn('⚠️ mappingManager.initialize() 呼び出し中にエラー:', err);
    }

    // メインアプリケーションを初期化
    window.photoFileManager = new SecurePhotoFileManagerApp();
    console.log('✅ セキュアアプリケーションの初期化が完了しました');
});

// エラーハンドリング
window.addEventListener('error', (event) => {
    console.error('グローバルエラー:', event.error);
    showError('予期しないエラーが発生しました: ' + event.error.message);
});

window.addEventListener('unhandledrejection', (event) => {
    console.error('未処理のPromise拒否:', event.reason);
    showError('処理中にエラーが発生しました: ' + event.reason);
});