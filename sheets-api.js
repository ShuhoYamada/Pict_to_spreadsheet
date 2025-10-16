// Google Sheets API管理クラス
class SheetsAPIManager {
    constructor() {
        this.selectedSpreadsheet = null;
        this.spreadsheetData = null;
    }

    // スプレッドシート一覧を取得
    async getSpreadsheets() {
        try {
            const response = await gapi.client.drive.files.list({
                q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
                pageSize: 100,
                fields: 'files(id,name,modifiedTime,owners)',
                orderBy: 'modifiedTime desc'
            });

            return response.result.files || [];
        } catch (error) {
            console.error('スプレッドシート取得エラー:', error);
            throw new Error('スプレッドシートの取得に失敗しました: ' + error.message);
        }
    }

    // 指定したスプレッドシートの情報を取得
    async getSpreadsheetInfo(spreadsheetId) {
        try {
            const response = await gapi.client.sheets.spreadsheets.get({
                spreadsheetId: spreadsheetId,
                fields: 'properties,sheets(properties(title,sheetId,gridProperties))'
            });

            return response.result;
        } catch (error) {
            console.error('スプレッドシート情報取得エラー:', error);
            throw new Error('スプレッドシート情報の取得に失敗しました: ' + error.message);
        }
    }

    // スプレッドシートの既存データを確認
    async checkExistingData(spreadsheetId, sheetName = null) {
        try {
            // シート名が指定されていない場合は最初のシートを使用
            const spreadsheetInfo = await this.getSpreadsheetInfo(spreadsheetId);
            const targetSheet = sheetName || spreadsheetInfo.sheets[0].properties.title;
            
            // A列からG列の2行目以降のデータを取得
            const range = `${targetSheet}!A2:G1000`;
            const response = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: range,
                valueRenderOption: 'UNFORMATTED_VALUE'
            });

            const values = response.result.values || [];
            
            // 各列の最後のデータがある行を特定
            const lastRows = {
                partName: this.findLastDataRow(values, 0),      // A列
                weight: this.findLastDataRow(values, 1),        // B列
                unit: this.findLastDataRow(values, 2),          // C列
                materialId: this.findLastDataRow(values, 3),    // D列
                processId: this.findLastDataRow(values, 4),     // E列
                photoType: this.findLastDataRow(values, 5),     // F列
                notes: this.findLastDataRow(values, 6)          // G列
            };

            return {
                sheetName: targetSheet,
                existingData: values,
                lastRows: lastRows,
                nextRows: {
                    partName: lastRows.partName + 2,      // 2行目から開始なので+2
                    weight: lastRows.weight + 2,
                    unit: lastRows.unit + 2,
                    materialId: lastRows.materialId + 2,
                    processId: lastRows.processId + 2,
                    photoType: lastRows.photoType + 2,
                    notes: lastRows.notes + 2
                }
            };
        } catch (error) {
            console.error('既存データ確認エラー:', error);
            throw new Error('既存データの確認に失敗しました: ' + error.message);
        }
    }

    // 指定した列で最後にデータがある行を検索
    findLastDataRow(values, columnIndex) {
        for (let i = values.length - 1; i >= 0; i--) {
            if (values[i] && values[i][columnIndex] && 
                values[i][columnIndex].toString().trim() !== '') {
                return i;
            }
        }
        return -1; // データが見つからない場合
    }

    // スプレッドシート選択モーダルを表示
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
                <div class="spreadsheet-item" onclick="sheetsManager.selectSpreadsheet('${sheet.id}', '${this.escapeHtml(sheet.name)}')">
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

    // スプレッドシートを選択
    async selectSpreadsheet(spreadsheetId, spreadsheetName) {
        try {
            showProgress('スプレッドシート情報を取得中...', 0);
            
            this.selectedSpreadsheet = {
                id: spreadsheetId,
                name: spreadsheetName
            };

            // スプレッドシート情報と既存データを取得
            const [spreadsheetInfo, existingDataInfo] = await Promise.all([
                this.getSpreadsheetInfo(spreadsheetId),
                this.checkExistingData(spreadsheetId)
            ]);

            this.spreadsheetData = {
                info: spreadsheetInfo,
                existing: existingDataInfo
            };
            
            hideProgress();
            this.updateSpreadsheetUI();
            document.getElementById('spreadsheet-modal').style.display = 'none';
            
            // 次のステップを有効化
            this.checkProcessAvailability();
            
        } catch (error) {
            hideProgress();
            showError(error.message);
        }
    }

    // スプレッドシート選択UIの更新
    updateSpreadsheetUI() {
        const spreadsheetInfo = document.getElementById('selected-spreadsheet-info');

        if (this.selectedSpreadsheet && this.spreadsheetData) {
            const existing = this.spreadsheetData.existing;
            spreadsheetInfo.className = 'info-box active';
            spreadsheetInfo.innerHTML = `
                <h4>✅ 選択されたスプレッドシート</h4>
                <p><strong>名前:</strong> ${this.escapeHtml(this.selectedSpreadsheet.name)}</p>
                <p><strong>シート:</strong> ${existing.sheetName}</p>
                <p><strong>既存データ:</strong> ${existing.existingData.length} 行</p>
                <div style="margin-top: 10px; font-size: 0.9rem;">
                    <strong>次回書き込み位置:</strong><br>
                    部品名: A${existing.nextRows.partName}行, 
                    重量: B${existing.nextRows.weight}行, 
                    単位: C${existing.nextRows.unit}行<br>
                    素材ID: D${existing.nextRows.materialId}行, 
                    加工ID: E${existing.nextRows.processId}行, 
                    写真区分: F${existing.nextRows.photoType}行, 
                    特記事項: G${existing.nextRows.notes}行
                </div>
            `;
        }
    }

    // データ処理の可用性をチェック
    checkProcessAvailability() {
        const photoFiles = driveManager.getSelectedPhotoFiles();
        if (this.selectedSpreadsheet && photoFiles.length > 0) {
            document.getElementById('process-data-button').disabled = false;
        }
    }

    // データをスプレッドシートに書き込み
    async writeDataToSpreadsheet(parsedDataArray) {
        if (!this.selectedSpreadsheet || !this.spreadsheetData) {
            throw new Error('スプレッドシートが選択されていません');
        }

        try {
            const spreadsheetId = this.selectedSpreadsheet.id;
            const sheetName = this.spreadsheetData.existing.sheetName;
            const existing = this.spreadsheetData.existing;
            
            const requests = [];
            
            // 各列に対してデータを書き込み
            for (let i = 0; i < parsedDataArray.length; i++) {
                const data = parsedDataArray[i];
                if (!data.isValid) continue;
                
                // 各列の次の行を計算
                const rowIndex = i;
                
                // A列: 部品名
                requests.push({
                    range: `${sheetName}!A${existing.nextRows.partName + rowIndex}`,
                    values: [[data.partName]]
                });
                
                // B列: 重量
                requests.push({
                    range: `${sheetName}!B${existing.nextRows.weight + rowIndex}`,
                    values: [[data.weightString]]
                });
                
                // C列: 単位
                requests.push({
                    range: `${sheetName}!C${existing.nextRows.unit + rowIndex}`,
                    values: [[data.unit]]
                });
                
                // D列: 素材ID
                requests.push({
                    range: `${sheetName}!D${existing.nextRows.materialId + rowIndex}`,
                    values: [[data.materialId]]
                });
                
                // E列: 加工ID
                requests.push({
                    range: `${sheetName}!E${existing.nextRows.processId + rowIndex}`,
                    values: [[data.processId]]
                });
                
                // F列: 写真区分
                requests.push({
                    range: `${sheetName}!F${existing.nextRows.photoType + rowIndex}`,
                    values: [[data.photoType]]
                });
                
                // G列: 特記事項
                if (data.notes) {
                    requests.push({
                        range: `${sheetName}!G${existing.nextRows.notes + rowIndex}`,
                        values: [[data.notes]]
                    });
                }
            }

            // バッチ更新を実行
            const batchUpdateRequest = {
                spreadsheetId: spreadsheetId,
                resource: {
                    valueInputOption: 'RAW',
                    data: requests
                }
            };

            const response = await gapi.client.sheets.spreadsheets.values.batchUpdate(batchUpdateRequest);
            
            return {
                success: true,
                updatedCells: response.result.totalUpdatedCells,
                updatedRows: response.result.totalUpdatedRows,
                requests: requests.length
            };

        } catch (error) {
            console.error('スプレッドシート書き込みエラー:', error);
            throw new Error('データの書き込みに失敗しました: ' + error.message);
        }
    }

    // HTMLエスケープ
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // 選択されたスプレッドシート情報を取得
    getSelectedSpreadsheet() {
        return this.selectedSpreadsheet;
    }

    // スプレッドシートデータを取得
    getSpreadsheetData() {
        return this.spreadsheetData;
    }
}

// スプレッドシートモーダルを閉じる
function closeSpreadsheetModal() {
    document.getElementById('spreadsheet-modal').style.display = 'none';
}

// Google Sheets APIマネージャーのインスタンス化
const sheetsManager = new SheetsAPIManager();