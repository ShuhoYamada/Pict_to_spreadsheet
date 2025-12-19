// セキュアな認証システム - バックエンド経由
class SecureAuthManager {
    constructor() {
        this.isAuthenticated = false;
        this.accessToken = null;
        this.init();
    }

    async init() {
        try {
            console.log('セキュア認証システムを初期化中...');
            this.setupEventListeners();
            this.checkAuthStatus();
        } catch (error) {
            console.error('認証システムの初期化に失敗しました:', error);
            showError('認証システムの初期化に失敗しました: ' + error.message);
        }
    }

    setupEventListeners() {
        const authorizeButton = document.getElementById('authorize-button');
        const signoutButton = document.getElementById('signout-button');
        
        if (authorizeButton) {
            authorizeButton.addEventListener('click', () => {
                this.signIn();
            });
            console.log('✅ 認証ボタンのイベントリスナーを登録しました');
        } else {
            console.error('❌ authorize-button が見つかりません');
        }

        if (signoutButton) {
            signoutButton.addEventListener('click', () => {
                this.signOut();
            });
            console.log('✅ サインアウトボタンのイベントリスナーを登録しました');
        } else {
            console.error('❌ signout-button が見つかりません');
        }
        
        // URLパラメータから認証結果をチェック
        const urlParams = new URLSearchParams(window.location.search);
        if (urlParams.get('auth') === 'success') {
            this.handleAuthSuccess();
        } else if (urlParams.get('auth') === 'error') {
            const errorMessage = urlParams.get('message') || '認証に失敗しました';
            console.error('🔴 認証エラー:', errorMessage);
            showError(`認証エラー: ${errorMessage}`);
        }
    }

    checkAuthStatus() {
        // セッションストレージから認証状態を確認
        const authData = sessionStorage.getItem('auth_data');
        if (authData) {
            try {
                const { isAuthenticated, timestamp } = JSON.parse(authData);
                const now = new Date().getTime();
                const oneHour = 60 * 60 * 1000;
                
                if (isAuthenticated && (now - timestamp) < oneHour) {
                    this.isAuthenticated = true;
                    this.updateAuthUI(true);
                    return;
                }
            } catch (e) {
                sessionStorage.removeItem('auth_data');
            }
        }
        
        this.updateAuthUI(false);
    }

    async signIn() {
        try {
            showStatus('認証URLを取得中...', 'info');
            
            const response = await fetch(`${CONFIG.API_BASE_URL}/auth/url`);
            const { authUrl } = await response.json();
            
            // 新しいウィンドウで認証ページを開く
            window.location.href = authUrl;
            
        } catch (error) {
            console.error('サインインエラー:', error);
            showError('認証の開始に失敗しました: ' + error.message);
        }
    }

    handleAuthSuccess() {
        this.isAuthenticated = true;
        
        // セッションストレージに認証状態を保存
        const authData = {
            isAuthenticated: true,
            timestamp: new Date().getTime()
        };
        sessionStorage.setItem('auth_data', JSON.stringify(authData));
        
        this.updateAuthUI(true);
        showStatus('認証が完了しました', 'success');
        
        // URLをクリーンアップ
        window.history.replaceState({}, document.title, window.location.pathname);
    }

    signOut() {
        this.isAuthenticated = false;
        this.accessToken = null;
        
        // セッションストレージをクリア
        sessionStorage.removeItem('auth_data');
        
        this.updateAuthUI(false);
        showStatus('サインアウトしました', 'info');
    }

    updateAuthUI(isAuthenticated) {
        const authorizeButton = document.getElementById('authorize-button');
        const signoutButton = document.getElementById('signout-button');
        const mainSection = document.getElementById('main-section');
        const authStatus = document.getElementById('auth-status');

        if (isAuthenticated) {
            authorizeButton.style.display = 'none';
            signoutButton.style.display = 'inline-block';
            mainSection.style.display = 'block';
            
            authStatus.innerHTML = `
                <div class="status-success">
                    ✅ 認証済み - セキュアな接続でGoogleサービスにアクセス中
                </div>
            `;
        } else {
            authorizeButton.style.display = 'inline-block';
            signoutButton.style.display = 'none';
            mainSection.style.display = 'none';
            authStatus.innerHTML = `
                <div class="status-info">
                    ℹ️ セキュアな認証が必要です
                </div>
            `;
        }
    }

    isUserAuthenticated() {
        return this.isAuthenticated;
    }
}

// グローバル関数
function showStatus(message, type = 'info') {
    const authStatus = document.getElementById('auth-status');
    const className = `status-${type}`;
    authStatus.innerHTML = `<div class="${className}">${message}</div>`;
}

function showError(message) {
    const errorContainer = document.getElementById('error-container');
    const errorMessage = document.getElementById('error-message');
    errorMessage.textContent = message;
    errorContainer.style.display = 'block';
    
    setTimeout(() => {
        hideError();
    }, 10000);
}

function hideError() {
    const errorContainer = document.getElementById('error-container');
    errorContainer.style.display = 'none';
}

function showMessage(message, type = 'info') {
    // 成功メッセージやその他の情報メッセージを表示
    console.log(`[${type.toUpperCase()}] ${message}`);
    
    // ステータス表示があれば使用
    if (typeof showStatus === 'function') {
        showStatus(message, type);
    }
}

// セキュア認証マネージャーのグローバル変数（DOMContentLoadedで初期化）
let authManager = null;

