// Google認証システム
class AuthManager {
    constructor() {
        this.isAuthenticated = false;
        this.authInstance = null;
        this.user = null;
        this.init();
    }

    async init() {
        try {
            // Google APIを初期化
            await this.loadGoogleAPI();
            await this.initGoogleAuth();
            this.setupEventListeners();
        } catch (error) {
            console.error('認証システムの初期化に失敗しました:', error);
            showError('認証システムの初期化に失敗しました: ' + error.message);
        }
    }

    async loadGoogleAPI() {
        return new Promise((resolve, reject) => {
            if (typeof gapi !== 'undefined') {
                resolve();
                return;
            }
            
            const script = document.createElement('script');
            script.src = 'https://apis.google.com/js/api.js';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    async initGoogleAuth() {
        if (!validateConfig()) {
            return;
        }

        return new Promise((resolve, reject) => {
            gapi.load('auth2', () => {
                gapi.auth2.init({
                    client_id: CONFIG.CLIENT_ID,
                    scope: CONFIG.SCOPES.join(' ')
                }).then(() => {
                    this.authInstance = gapi.auth2.getAuthInstance();
                    this.checkAuthStatus();
                    resolve();
                }).catch(reject);
            });
        });
    }

    setupEventListeners() {
        document.getElementById('authorize-button').addEventListener('click', () => {
            this.signIn();
        });

        document.getElementById('signout-button').addEventListener('click', () => {
            this.signOut();
        });
    }

    checkAuthStatus() {
        if (this.authInstance) {
            const isSignedIn = this.authInstance.isSignedIn.get();
            this.updateAuthUI(isSignedIn);
            
            if (isSignedIn) {
                this.user = this.authInstance.currentUser.get();
                this.isAuthenticated = true;
                this.loadGoogleAPIs();
            }
        }
    }

    async signIn() {
        try {
            showStatus('認証中...', 'info');
            const user = await this.authInstance.signIn();
            this.user = user;
            this.isAuthenticated = true;
            this.updateAuthUI(true);
            await this.loadGoogleAPIs();
            showStatus('認証が完了しました', 'success');
        } catch (error) {
            console.error('サインインエラー:', error);
            showError('認証に失敗しました: ' + error.error);
        }
    }

    async signOut() {
        try {
            await this.authInstance.signOut();
            this.user = null;
            this.isAuthenticated = false;
            this.updateAuthUI(false);
            showStatus('サインアウトしました', 'info');
        } catch (error) {
            console.error('サインアウトエラー:', error);
            showError('サインアウトに失敗しました');
        }
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
            
            if (this.user) {
                const profile = this.user.getBasicProfile();
                authStatus.innerHTML = `
                    <div class="status-success">
                        ✅ ${profile.getName()} (${profile.getEmail()}) でログイン中
                    </div>
                `;
            }
        } else {
            authorizeButton.style.display = 'inline-block';
            signoutButton.style.display = 'none';
            mainSection.style.display = 'none';
            authStatus.innerHTML = `
                <div class="status-info">
                    ℹ️ Googleアカウントで認証してください
                </div>
            `;
        }
    }

    async loadGoogleAPIs() {
        try {
            // Google Drive APIを読み込み
            await new Promise((resolve, reject) => {
                gapi.client.load('drive', 'v3').then(resolve).catch(reject);
            });

            // Google Sheets APIを読み込み
            await new Promise((resolve, reject) => {
                gapi.client.load('sheets', 'v4').then(resolve).catch(reject);
            });

            console.log('Google APIs loaded successfully');
        } catch (error) {
            console.error('Google APIs読み込みエラー:', error);
            showError('Google APIsの読み込みに失敗しました');
        }
    }

    getAccessToken() {
        if (this.user && this.isAuthenticated) {
            return this.user.getAuthResponse().access_token;
        }
        return null;
    }

    isUserAuthenticated() {
        return this.isAuthenticated && this.user;
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
    
    // 10秒後に自動で閉じる
    setTimeout(() => {
        hideError();
    }, 10000);
}

function hideError() {
    const errorContainer = document.getElementById('error-container');
    errorContainer.style.display = 'none';
}

// 認証マネージャーのインスタンス化
let authManager;