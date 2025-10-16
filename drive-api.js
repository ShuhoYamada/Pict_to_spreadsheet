// Google Drive API管理クラス
class DriveAPIManager {
    constructor() {
        this.selectedFolder = null;
        this.photoFiles = [];
    }

    // フォルダ一覧を取得
    async getFolders() {
        try {
            const response = await gapi.client.drive.files.list({
                q: "mimeType='application/vnd.google-apps.folder' and trashed=false",
                pageSize: 100,
                fields: 'files(id,name,parents,modifiedTime)',
                orderBy: 'name'
            });

            return response.result.files || [];
        } catch (error) {
            console.error('フォルダ取得エラー:', error);
            throw new Error('フォルダの取得に失敗しました: ' + error.message);
        }
    }

    // 指定フォルダ内の写真ファイルを取得
    async getPhotosInFolder(folderId) {
        try {
            const photoExtensions = CONFIG.PHOTO_EXTENSIONS.join('|');
            const query = `'${folderId}' in parents and trashed=false and (${CONFIG.PHOTO_EXTENSIONS.map(ext => `name contains '.${ext}'`).join(' or ')})`;
            
            const response = await gapi.client.drive.files.list({
                q: query,
                pageSize: 1000,
                fields: 'files(id,name,size,modifiedTime)',
                orderBy: 'name'
            });

            const files = response.result.files || [];
            
            // 拡張子を再確認してフィルタリング
            return files.filter(file => {
                const fileName = file.name.toLowerCase();
                return CONFIG.PHOTO_EXTENSIONS.some(ext => fileName.endsWith(`.${ext}`));
            });
        } catch (error) {
            console.error('写真ファイル取得エラー:', error);
            throw new Error('写真ファイルの取得に失敗しました: ' + error.message);
        }
    }

    // フォルダ選択モーダルを表示
    async showFolderSelector() {
        if (!authManager.isUserAuthenticated()) {
            showError('認証が必要です');
            return;
        }

        try {
            showProgress('フォルダを読み込み中...', 0);
            const folders = await this.getFolders();
            hideProgress();

            this.renderFolderList(folders);
            document.getElementById('folder-modal').style.display = 'flex';
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // フォルダリストをレンダリング
    renderFolderList(folders) {
        const folderList = document.getElementById('folder-list');
        
        if (folders.length === 0) {
            folderList.innerHTML = '<p>フォルダが見つかりませんでした。</p>';
            return;
        }

        folderList.innerHTML = folders.map(folder => `
            <div class="folder-item" onclick="driveManager.selectFolder('${folder.id}', '${folder.name}')">
                <span>📁</span>
                <div>
                    <strong>${this.escapeHtml(folder.name)}</strong>
                    <div style="font-size: 0.8rem; color: #666; margin-top: 4px;">
                        最終更新: ${new Date(folder.modifiedTime).toLocaleDateString('ja-JP')}
                    </div>
                </div>
            </div>
        `).join('');
    }

    // フォルダを選択
    async selectFolder(folderId, folderName) {
        try {
            showProgress('写真ファイルを確認中...', 0);
            
            this.selectedFolder = {
                id: folderId,
                name: folderName
            };

            // 写真ファイルを取得
            this.photoFiles = await this.getPhotosInFolder(folderId);
            
            hideProgress();
            this.updateFolderUI();
            document.getElementById('folder-modal').style.display = 'none';
            
            // 次のステップを有効化
            this.checkNextStepAvailability();
            
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // フォルダ選択UIの更新
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

    // ファイルプレビューをレンダリング
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

    // 次のステップの可用性をチェック
    checkNextStepAvailability() {
        // フォルダが選択され、写真ファイルがある場合のみ次のステップを有効化
        if (this.selectedFolder && this.photoFiles.length > 0) {
            // スプレッドシート選択ボタンを有効化
            document.getElementById('select-spreadsheet-button').disabled = false;
        }
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // 選択されたフォルダの写真ファイルを取得
    getSelectedPhotoFiles() {
        return this.photoFiles || [];
    }

    // 選択されたフォルダ情報を取得
    getSelectedFolder() {
        return this.selectedFolder;
    }
}

// フォルダモーダルを閉じる
function closeFolderModal() {
    document.getElementById('folder-modal').style.display = 'none';
}

// プログレス表示
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

// Google Drive APIマネージャーのインスタンス化
const driveManager = new DriveAPIManager();