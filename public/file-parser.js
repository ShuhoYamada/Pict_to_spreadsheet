// ファイル名解析クラス - 新仕様対応
class FileNameParser {
    constructor() {
        // 仕様: 番号_部品名_重量_単位_素材ID_加工ID_写真区分_特記事項.拡張子
        // 例: 001_ねじ_10.1_g_st01_pr01_P_0.jpg
        // 番号は無視して処理、残りの7項目を使用
        // 特記事項は必須項目
        // 写真区分がMの場合は処理対象外、Pの場合のみ処理対象
    }

    // ファイル名を解析
    parseFileName(fileName) {
        try {
            // 拡張子を除去
            const nameWithoutExt = fileName.substring(0, fileName.lastIndexOf('.'));
            
            // アンダーバーで分割
            const parts = nameWithoutExt.split('_');
            
            // 必須項目: 番号、部品名、重量、単位、素材ID、加工ID、写真区分、特記事項（8項目）
            // 番号は無視するが、形式チェックのため8項目必要
            if (parts.length < 8) {
                return {
                    isValid: false,
                    error: `必要な項目が不足しています（8項目必要、現在${parts.length}項目）`,
                    fileName: fileName,
                    parts: parts,
                    requiredFormat: '番号_部品名_重量_単位_素材ID_加工ID_写真区分_特記事項.拡張子'
                };
            }

            // 各部分を解析（番号[0]は無視、[1]以降を使用）
            const number = parts[0];      // 番号（処理では無視）
            const partName = parts[1];    // 部品名
            const weight = parts[2];      // 重量
            const unit = parts[3];        // 単位
            const materialId = parts[4];  // 素材ID
            const processId = parts[5];   // 加工ID
            const photoType = parts[6];   // 写真区分
            const notes = parts[7];       // 特記事項

            // 写真区分の妥当性チェック
            if (!photoType || (photoType.toLowerCase() !== 'p' && photoType.toLowerCase() !== 'm')) {
                return {
                    isValid: false,
                    error: '写真区分は「P」または「M」である必要があります',
                    fileName: fileName,
                    parts: parts,
                    photoType: photoType
                };
            }

            // 写真区分がMの場合は処理対象外
            if (photoType.toLowerCase() === 'm') {
                return {
                    isValid: true,
                    shouldSkip: true,
                    skipReason: '写真区分がMのため処理対象外',
                    fileName: fileName,
                    partName: partName,
                    photoType: photoType
                };
            }

            // 特記事項のチェック（必須項目）
            if (notes === undefined || notes === '') {
                return {
                    isValid: false,
                    error: '特記事項は必須項目です',
                    fileName: fileName,
                    parts: parts
                };
            }

            // 重量の妥当性チェック
            const weightNum = parseFloat(weight);
            if (isNaN(weightNum) || weightNum < 0) {
                return {
                    isValid: false,
                    error: '重量が数値として認識できません',
                    fileName: fileName,
                    parts: parts
                };
            }

            // 単位の妥当性チェック（サポートする単位）
            const validUnits = ['g', 'kg'];
            if (!validUnits.includes(unit.toLowerCase())) {
                return {
                    isValid: false,
                    error: `単位は「g」または「kg」である必要があります（現在: ${unit}）`,
                    fileName: fileName,
                    parts: parts
                };
            }

            // 特記事項の値をチェック（0または1）
            if (notes !== '0' && notes !== '1') {
                return {
                    isValid: false,
                    error: '特記事項は「0」または「1」である必要があります',
                    fileName: fileName,
                    parts: parts,
                    notes: notes
                };
            }

            return {
                isValid: true,
                shouldSkip: false,
                fileName: fileName,
                partName: partName,
                weight: weightNum,
                weightString: weight, // 元の文字列も保持
                unit: unit.toLowerCase(),
                materialId: materialId,
                processId: processId,
                photoType: photoType.toLowerCase(),
                notes: notes,
                parts: parts,
                parsedAt: new Date(),
                // 単位変換用の重量値を事前計算
                weightInGrams: unit.toLowerCase() === 'kg' ? weightNum * 1000 : weightNum,
                weightInKilograms: unit.toLowerCase() === 'g' ? weightNum / 1000 : weightNum,
                // 特記事項の変換
                notesText: notes === '0' ? '-' : 'あり'
            };

        } catch (error) {
            console.error('ファイル名解析エラー:', error);
            return {
                isValid: false,
                error: 'ファイル名の解析中にエラーが発生しました: ' + error.message,
                fileName: fileName
            };
        }
    }

    // 複数のファイルを一括解析
    parseMultipleFiles(files) {
        const results = {
            valid: [],
            invalid: [],
            summary: {
                total: files.length,
                validCount: 0,
                invalidCount: 0
            }
        };

        files.forEach(file => {
            const parsed = this.parseFileName(file.name);
            parsed.fileInfo = file; // 元のファイル情報も保持
            
            if (parsed.isValid) {
                results.valid.push(parsed);
                results.summary.validCount++;
            } else {
                results.invalid.push(parsed);
                results.summary.invalidCount++;
            }
        });

        return results;
    }

    // 解析結果を検証
    validateParsedData(parsedData) {
        const errors = [];
        const warnings = [];

        // 部品名の検証
        if (!parsedData.partName || parsedData.partName.trim().length === 0) {
            errors.push('部品名が空です');
        }

        // 重量の検証
        if (parsedData.weight <= 0) {
            errors.push('重量は正の数値である必要があります');
        }

        // ID形式の検証（英数字とハイフンのみ許可）
        const idPattern = /^[a-zA-Z0-9\-]+$/;
        if (!idPattern.test(parsedData.materialId)) {
            warnings.push(`素材ID "${parsedData.materialId}" に特殊文字が含まれています`);
        }
        if (!idPattern.test(parsedData.processId)) {
            warnings.push(`加工ID "${parsedData.processId}" に特殊文字が含まれています`);
        }

        return {
            isValid: errors.length === 0,
            errors: errors,
            warnings: warnings
        };
    }

    // スプレッドシート用のデータ形式に変換
    formatForSpreadsheet(parsedDataArray) {
        return parsedDataArray.map(data => {
            if (!data.isValid) return null;
            
            return [
                data.partName,           // A列: 部品名
                data.weightString,       // B列: 重量（元の文字列形式）
                data.unit,              // C列: 単位
                data.materialId,        // D列: 素材ID
                data.processId,         // E列: 加工ID
                data.photoType,         // F列: 写真区分
                data.notes || ''        // G列: 特記事項
            ];
        }).filter(row => row !== null);
    }

    // 解析統計を生成
    generateStatistics(parseResults) {
        const stats = {
            totalFiles: parseResults.summary.total,
            validFiles: parseResults.summary.validCount,
            invalidFiles: parseResults.summary.invalidCount,
            successRate: parseResults.summary.total > 0 ? 
                ((parseResults.summary.validCount / parseResults.summary.total) * 100).toFixed(1) : 0,
            partNames: new Set(),
            units: new Set(),
            materialIds: new Set(),
            processIds: new Set(),
            photoTypes: new Set(),
            errors: []
        };

        // 有効なファイルから統計を収集
        parseResults.valid.forEach(data => {
            stats.partNames.add(data.partName);
            stats.units.add(data.unit);
            stats.materialIds.add(data.materialId);
            stats.processIds.add(data.processId);
            stats.photoTypes.add(data.photoType);
        });

        // 無効なファイルのエラーを収集
        parseResults.invalid.forEach(data => {
            stats.errors.push({
                fileName: data.fileName,
                error: data.error
            });
        });

        // SetをArrayに変換
        stats.partNames = Array.from(stats.partNames);
        stats.units = Array.from(stats.units);
        stats.materialIds = Array.from(stats.materialIds);
        stats.processIds = Array.from(stats.processIds);
        stats.photoTypes = Array.from(stats.photoTypes);

        return stats;
    }

    // サンプルファイル名を生成（テスト用）
    generateSampleFileName() {
        const samples = [
            'ねじ_10.1_g_ol01_pr01_重量.jpg',
            'ワッシャー_5.2_g_st02_nc01_外観.png',
            'ナット_15.0_g_al03_mc02_断面_耐食性.jpeg',
            'ボルト_25.8_g_ti04_la03_全体.heic'
        ];
        return samples[Math.floor(Math.random() * samples.length)];
    }
}

// ファイル名パーサーのインスタンス化
const fileParser = new FileNameParser();