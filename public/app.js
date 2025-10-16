// メインアプリケーション
class PhotoFileManagerApp {
    constructor() {
        this.isInitialized = false;
        this.processingData = null;
        this.init();
    }

    // アプリケーション初期化
    async init() {
        try {
            console.log('写真ファイル管理システムを初期化中...');
            
            // 設定を検証
            if (!validateConfig()) {
                return;
            }

            // 認証マネージャーを初期化
            authManager = new AuthManager();
            
            // イベントリスナーを設定
            this.setupEventListeners();
            
            this.isInitialized = true;
            console.log('アプリケーションの初期化が完了しました');
            
        } catch (error) {
            console.error('アプリケーション初期化エラー:', error);
            showError('アプリケーションの初期化に失敗しました: ' + error.message);
        }
    }

    // イベントリスナーを設定
    setupEventListeners() {
        // フォルダ選択ボタン
        document.getElementById('select-folder-button').addEventListener('click', () => {
            driveManager.showFolderSelector();
        });

        // スプレッドシート選択ボタン
        document.getElementById('select-spreadsheet-button').addEventListener('click', () => {
            sheetsManager.showSpreadsheetSelector();
        });

        // データ処理ボタン
        document.getElementById('process-data-button').addEventListener('click', () => {
            this.processAndWriteData();
        });

        // ESCキーでモーダルを閉じる
        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                document.getElementById('folder-modal').style.display = 'none';
                document.getElementById('spreadsheet-modal').style.display = 'none';
                hideError();
            }
        });

        // モーダル背景クリックで閉じる
        document.getElementById('folder-modal').addEventListener('click', (event) => {
            if (event.target === event.currentTarget) {
                closeFolderModal();
            }
        });

        document.getElementById('spreadsheet-modal').addEventListener('click', (event) => {
            if (event.target === event.currentTarget) {
                closeSpreadsheetModal();
            }
        });
    }

    // データを処理してスプレッドシートに書き込み
    async processAndWriteData() {
        try {
            if (!authManager.isUserAuthenticated()) {
                showError('認証が必要です');
                return;
            }

            const photoFiles = driveManager.getSelectedPhotoFiles();
            const selectedFolder = driveManager.getSelectedFolder();
            const selectedSpreadsheet = sheetsManager.getSelectedSpreadsheet();

            if (!photoFiles || photoFiles.length === 0) {
                showError('処理する写真ファイルがありません');
                return;
            }

            if (!selectedSpreadsheet) {
                showError('スプレッドシートが選択されていません');
                return;
            }

            // 処理開始
            this.updateProcessingStatus('📋 ファイル名を解析中...', 10);
            
            // ファイル名を解析
            const parseResults = fileParser.parseMultipleFiles(photoFiles);
            
            if (parseResults.summary.validCount === 0) {
                throw new Error('有効なファイル名形式の写真ファイルが見つかりませんでした');
            }

            this.updateProcessingStatus('📊 解析結果を検証中...', 30);

            // 解析結果を検証
            const validData = parseResults.valid;
            const invalidData = parseResults.invalid;

            // 統計情報を生成
            const statistics = fileParser.generateStatistics(parseResults);

            this.updateProcessingStatus('💾 スプレッドシートに書き込み中...', 60);

            // スプレッドシートに書き込み
            const writeResult = await sheetsManager.writeDataToSpreadsheet(validData);

            this.updateProcessingStatus('✅ 処理完了!', 100);

            // 結果を表示
            this.showResults({
                folder: selectedFolder,
                spreadsheet: selectedSpreadsheet,
                parseResults: parseResults,
                statistics: statistics,
                writeResult: writeResult,
                validData: validData,
                invalidData: invalidData
            });

            // 処理完了後の状態更新
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

    // 処理状況を更新
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

    // 結果を表示
    showResults(results) {
        const resultsBox = document.getElementById('results');
        const { folder, spreadsheet, parseResults, statistics, writeResult, validData, invalidData } = results;

        let resultsHTML = `
            <h4>🎉 処理結果</h4>
            
            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📂 処理したフォルダ</h5>
                <p><strong>${this.escapeHtml(folder.name)}</strong></p>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📊 書き込み先スプレッドシート</h5>
                <p><strong>${this.escapeHtml(spreadsheet.name)}</strong></p>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>📈 統計情報</h5>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
                    <div class="stat-item">
                        <strong>総ファイル数:</strong> ${statistics.totalFiles}個
                    </div>
                    <div class="stat-item">
                        <strong>処理成功:</strong> ${statistics.validFiles}個
                    </div>
                    <div class="stat-item">
                        <strong>処理失敗:</strong> ${statistics.invalidFiles}個
                    </div>
                    <div class="stat-item">
                        <strong>成功率:</strong> ${statistics.successRate}%
                    </div>
                </div>
            </div>

            <div class="result-section" style="margin-bottom: 20px;">
                <h5>💾 書き込み結果</h5>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
                    <div class="stat-item">
                        <strong>更新セル数:</strong> ${writeResult.updatedCells}個
                    </div>
                    <div class="stat-item">
                        <strong>更新行数:</strong> ${writeResult.updatedRows}行
                    </div>
                </div>
            </div>
        `;

        // 有効なデータの詳細
        if (validData.length > 0) {
            resultsHTML += `
                <div class="result-section" style="margin-bottom: 20px;">
                    <h5>✅ 処理成功したファイル (${validData.length}個)</h5>
                    <div style="max-height: 200px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px;">
                        ${validData.slice(0, 10).map(data => `
                            <div style="margin-bottom: 8px; padding: 8px; background: #f0fff4; border-radius: 4px; border-left: 3px solid #48bb78;">
                                <strong>${this.escapeHtml(data.fileName)}</strong><br>
                                <small>部品: ${data.partName} | 重量: ${data.weightString}${data.unit} | 素材: ${data.materialId} | 加工: ${data.processId}</small>
                            </div>
                        `).join('')}
                        ${validData.length > 10 ? `<p><em>他 ${validData.length - 10} 件...</em></p>` : ''}
                    </div>
                </div>
            `;
        }

        // 無効なデータの詳細
        if (invalidData.length > 0) {
            resultsHTML += `
                <div class="result-section">
                    <h5>❌ 処理失敗したファイル (${invalidData.length}個)</h5>
                    <div style="max-height: 200px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px;">
                        ${invalidData.slice(0, 10).map(data => `
                            <div style="margin-bottom: 8px; padding: 8px; background: #fff5f5; border-radius: 4px; border-left: 3px solid #f56565;">
                                <strong>${this.escapeHtml(data.fileName)}</strong><br>
                                <small style="color: #e53e3e;">エラー: ${data.error}</small>
                            </div>
                        `).join('')}
                        ${invalidData.length > 10 ? `<p><em>他 ${invalidData.length - 10} 件...</em></p>` : ''}
                    </div>
                </div>
            `;
        }

        resultsBox.className = 'results-box active';
        resultsBox.innerHTML = resultsHTML;

        // 成功メッセージを表示
        const statusElement = document.getElementById('processing-status');
        statusElement.className = 'status-message status-success';
        statusElement.innerHTML = `
            ✅ データの処理が完了しました！ ${statistics.validFiles}個のファイルがスプレッドシートに書き込まれました。
        `;
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // アプリケーションの状態をリセット
    resetApplication() {
        driveManager.selectedFolder = null;
        driveManager.photoFiles = [];
        sheetsManager.selectedSpreadsheet = null;
        sheetsManager.spreadsheetData = null;

        // UIをリセット
        document.getElementById('selected-folder-info').className = 'info-box';
        document.getElementById('photo-files-preview').className = 'preview-box';
        document.getElementById('selected-spreadsheet-info').className = 'info-box';
        document.getElementById('results').className = 'results-box';
        document.getElementById('processing-status').innerHTML = '';
        
        // ボタンを無効化
        document.getElementById('select-spreadsheet-button').disabled = true;
        document.getElementById('process-data-button').disabled = true;
    }
}

// ページ読み込み完了後にアプリケーションを初期化
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM読み込み完了 - アプリケーションを初期化します');
    window.photoFileManager = new PhotoFileManagerApp();
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