# 📚 統合学習アプリ (Integrated Learning App)

英単語・例文学習とクイズ機能を統合したPythonアプリケーションです。

## ✨ 機能

### 📖 学習機能
- **単語管理**: 英単語と日本語意味の登録・編集・削除
- **例文管理**: 英語例文と日本語訳の登録・編集・削除
- **翻訳機能**: DeepL翻訳・Google翻訳対応

### 🎯 クイズ機能
- **4択クイズ**: 単語の意味を4つの選択肢から選択
- **リスニングクイズ**: 音声を聞いて意味を答える
- **例文クイズ**: 例文のリスニングクイズ
- **日付フィルタ**: 登録日で出題範囲を絞り込み

### 🔄 復習システム
- **間違い自動記録**: 間違えた問題を自動的に記録
- **3回連続正解で卒業**: 習得済み問題は復習対象から除外
- **専用復習クイズ**: 間違えた問題のみを出題

### 🎵 音声機能
- **TTS読み上げ**: edge-ttsによる英語音声合成
- **効果音**: 正解・不正解時の音声フィードバック

## 🚀 使用方法

### 開発環境での実行
```bash
# 1. 仮想環境作成
python -m venv .venv

# 2. 仮想環境アクティベート
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # Mac/Linux

# 3. 依存関係インストール
pip install deepl edge-tts tkcalendar deep-translator

# 4. アプリ実行
python integrated_learning_app.py
実行ファイル作成
bash# PyInstallerでexe作成（Windows）
make_exe.bat

# 配布用パッケージ作成
create_package.bat
📦 必要なライブラリ
bashpip install deepl edge-tts tkcalendar deep-translator
📁 ファイル構成
integrated-learning-app/
├── integrated_learning_app.py    # メインアプリケーション
├── make_exe.bat                  # exe作成用バッチファイル
├── create_package.bat            # 配布用パッケージ作成
├── start_app.bat                 # 開発用起動スクリプト
├── sound_correct.mp3             # 正解効果音
├── sound_incorrect.mp3           # 不正解効果音
├── README.md                     # このファイル
└── .gitignore                    # Git除外設定
⚙️ 設定
DeepL翻訳設定

アプリのメニューから「翻訳設定」を開く
DeepL APIでAPIキーを取得
APIキーを入力して保存

Google翻訳

APIキー不要で無料利用可能
deep-translatorライブラリ経由で利用

🎯 対応環境

OS: Windows 7/8/10/11
Python: 3.8以上
GUI: tkinter（Pythonに標準装備）

📊 データベース

SQLite3を使用してローカルにデータ保存
自動的にlearning.sqlite3ファイルが作成されます

🛠️ 開発者向け
ビルド
bash# PyInstallerでスタンドアロン実行ファイル作成
pyinstaller --onefile --windowed --name "LearningApp" integrated_learning_app.py
カスタマイズ

音声ファイルの変更: sound_correct.mp3, sound_incorrect.mp3を差し替え
UIの調整: integrated_learning_app.py内のtkinter設定を変更

📄 ライセンス
MIT License
🤝 貢献
Issues やプルリクエストは歓迎です！
📧 連絡先
何か質問があれば、GitHubのIssuesでお知らせください