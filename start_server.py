#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
写真ファイル管理システム サーバー起動スクリプト
VSCodeの実行ボタン（▶）でサーバーを起動できます
"""

import os
import sys
import subprocess
import time
import signal
import webbrowser
from pathlib import Path

class PhotoFileManagerServer:
    def __init__(self):
        self.server_process = None
        self.server_url = "http://localhost:3000"
        self.project_dir = Path(__file__).parent
        
    def check_dependencies(self):
        """依存関係をチェック"""
        print("🔍 依存関係をチェック中...")
        
        # Node.jsがインストールされているかチェック
        try:
            result = subprocess.run(['node', '--version'], 
                                  capture_output=True, text=True, cwd=self.project_dir)
            if result.returncode == 0:
                print(f"✅ Node.js: {result.stdout.strip()}")
            else:
                raise FileNotFoundError
        except FileNotFoundError:
            print("❌ Node.jsがインストールされていません！")
            print("   https://nodejs.org/ からNode.jsをインストールしてください")
            return False
        
        # package.jsonが存在するかチェック
        package_json = self.project_dir / "package.json"
        if not package_json.exists():
            print("❌ package.jsonが見つかりません！")
            return False
        print("✅ package.json: 存在確認")
        
        # .envファイルが存在するかチェック
        env_file = self.project_dir / ".env"
        if not env_file.exists():
            print("❌ .envファイルが見つかりません！")
            print("   Google Cloud Consoleの認証情報を設定してください")
            return False
        print("✅ .env: 存在確認")
        
        # node_modulesが存在するかチェック
        node_modules = self.project_dir / "node_modules"
        if not node_modules.exists():
            print("📦 依存関係をインストール中...")
            try:
                result = subprocess.run(['npm', 'install'], 
                                      cwd=self.project_dir, 
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    print("✅ 依存関係のインストール完了")
                else:
                    print(f"❌ 依存関係のインストールに失敗: {result.stderr}")
                    return False
            except Exception as e:
                print(f"❌ npm installでエラー: {e}")
                return False
        else:
            print("✅ node_modules: 存在確認")
        
        return True
    
    def check_port(self):
        """ポート3000が使用中かチェック"""
        try:
            result = subprocess.run(['lsof', '-ti:3000'], 
                                  capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                print("⚠️  ポート3000が使用中です。既存のプロセスを停止します...")
                subprocess.run(['kill', '-9'] + result.stdout.strip().split('\n'), 
                             capture_output=True)
                time.sleep(1)
                print("✅ ポート3000を解放しました")
        except Exception as e:
            print(f"ポートチェック中にエラー: {e}")
    
    def start_server(self):
        """サーバーを起動"""
        print("🚀 サーバーを起動中...")
        
        try:
            # サーバープロセスを起動
            self.server_process = subprocess.Popen(
                ['npm', 'start'],
                cwd=self.project_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )
            
            # サーバーの起動を監視
            server_started = False
            start_time = time.time()
            timeout = 30  # 30秒タイムアウト
            
            print("📡 サーバー起動を監視中...")
            
            while time.time() - start_time < timeout:
                if self.server_process.poll() is not None:
                    # プロセスが終了した場合
                    print("❌ サーバーの起動に失敗しました")
                    return False
                
                # サーバーの出力を読み取り
                try:
                    line = self.server_process.stdout.readline()
                    if line:
                        print(f"   {line.strip()}")
                        
                        # サーバー起動完了の確認
                        if "サーバーが起動しました" in line or "server started" in line.lower():
                            server_started = True
                            break
                        elif "🚀" in line and "localhost:3000" in line:
                            server_started = True
                            break
                            
                except Exception:
                    pass
                
                time.sleep(0.1)
            
            if not server_started:
                print("⏰ サーバー起動のタイムアウト（30秒）")
                return False
            
            print("✅ サーバーが正常に起動しました！")
            print(f"🌐 アクセスURL: {self.server_url}")
            return True
            
        except Exception as e:
            print(f"❌ サーバー起動エラー: {e}")
            return False
    
    def open_browser(self):
        """ブラウザを自動で開く"""
        try:
            print("🌐 ブラウザを開いています...")
            webbrowser.open(self.server_url)
            print(f"✅ ブラウザで {self.server_url} を開きました")
        except Exception as e:
            print(f"⚠️  ブラウザの自動起動に失敗: {e}")
            print(f"   手動でブラウザを開いて {self.server_url} にアクセスしてください")
    
    def setup_signal_handler(self):
        """シグナルハンドラーを設定（Ctrl+Cでの停止）"""
        def signal_handler(signum, frame):
            print("\n🛑 停止要求を受信しました...")
            self.stop_server()
            sys.exit(0)
        
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
    
    def stop_server(self):
        """サーバーを停止"""
        if self.server_process:
            print("🛑 サーバーを停止中...")
            try:
                self.server_process.terminate()
                self.server_process.wait(timeout=5)
                print("✅ サーバーを正常に停止しました")
            except subprocess.TimeoutExpired:
                print("⚠️  強制停止を実行中...")
                self.server_process.kill()
                print("✅ サーバーを強制停止しました")
            except Exception as e:
                print(f"❌ サーバー停止エラー: {e}")
    
    def monitor_server(self):
        """サーバーの動作を監視"""
        print("\n" + "="*60)
        print("📋 写真ファイル管理システム - サーバー監視中")
        print("="*60)
        print(f"🌐 URL: {self.server_url}")
        print("🎯 使用方法:")
        print("   1. Google認証を実行")
        print("   2. フォルダを選択してGoogleドライブの写真を取得")
        print("   3. スプレッドシートを選択してデータ書き込み先を設定")
        print("   4. データ処理を実行（写真ハイパーリンク自動設定）")
        print("🔗 新機能: 写真ハイパーリンク機能")
        print("   - P区分写真 → 構成部品列にリンク設定")
        print("   - M区分写真 → 対応するP区分行の素材列にリンク設定")
        print("\n💡 停止するには Ctrl+C を押してください")
        print("="*60)
        
        try:
            # サーバーの出力を継続的に表示
            while True:
                if self.server_process.poll() is not None:
                    print("❌ サーバープロセスが終了しました")
                    break
                
                try:
                    line = self.server_process.stdout.readline()
                    if line:
                        print(f"📡 {line.strip()}")
                except Exception:
                    pass
                
                time.sleep(0.1)
                
        except KeyboardInterrupt:
            print("\n🛑 キーボード割り込みを受信")
    
    def run(self):
        """メイン実行関数"""
        print("🚀 写真ファイル管理システム サーバー起動スクリプト")
        print("="*60)
        
        # シグナルハンドラーを設定
        self.setup_signal_handler()
        
        # 依存関係をチェック
        if not self.check_dependencies():
            print("\n❌ セットアップに失敗しました")
            input("Enterキーを押して終了...")
            return
        
        # ポートをチェック
        self.check_port()
        
        # サーバーを起動
        if not self.start_server():
            print("\n❌ サーバーの起動に失敗しました")
            input("Enterキーを押して終了...")
            return
        
        # ブラウザを開く
        time.sleep(2)  # サーバーの完全起動を待つ
        self.open_browser()
        
        # サーバーを監視
        try:
            self.monitor_server()
        except Exception as e:
            print(f"❌ 監視中にエラー: {e}")
        finally:
            self.stop_server()


def main():
    """メイン関数"""
    try:
        server_manager = PhotoFileManagerServer()
        server_manager.run()
    except Exception as e:
        print(f"❌ 予期しないエラー: {e}")
        input("Enterキーを押して終了...")


if __name__ == "__main__":
    main()