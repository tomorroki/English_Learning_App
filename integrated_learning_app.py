"""
integrated_learning_app.py (v3.6 - 単語・例文統合学習アプリ - Google翻訳ボタン追加版 完全版)
"""
import json, os, re, random, sqlite3, sys, tkinter as tk, shutil, subprocess
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
import traceback
import threading
import asyncio
import ctypes

# --- ライブラリのインポート ---
try:
    import deepl
    HAS_TRANSLATOR = True
except ImportError:
    deepl = None
    HAS_TRANSLATOR = False

try:
    import edge_tts
    HAS_TTS = True
except ImportError:
    edge_tts = None
    HAS_TTS = False

try:
    from tkcalendar import DateEntry
    HAS_CALENDAR = True
except ImportError:
    DateEntry = None
    HAS_CALENDAR = False

try:
    from deep_translator import GoogleTranslator
    HAS_GOOGLE_TRANSLATOR = True
except ImportError:
    GoogleTranslator = None
    HAS_GOOGLE_TRANSLATOR = False

# --- 定数定義 ---
APP_TITLE = "統合学習アプリ (単語・例文・クイズ) v3.6"
CONFIG_FILE = "integrated_config.json"
DEFAULT_DB_FILE = "learning.sqlite3"
AUDIO_FILE_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), "temp_speech.mp3")
TTS_VOICE = "en-US-JennyNeural"
SOUND_FILES = {
    "correct": "sound_correct.mp3",
    "incorrect": "sound_incorrect.mp3"
}

# --- データベース管理クラス ---
class DatabaseManager:
    def __init__(self, db_path):
        self.conn = sqlite3.connect(db_path)
        self.create_tables()
    
    def create_tables(self):
        with self.conn:
            # 例文テーブル
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS sentences (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    english_sentence TEXT NOT NULL UNIQUE,
                    japanese_translation TEXT NOT NULL,
                    created_at INTEGER
                )
            """)
            
            # 単語テーブル
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS words (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    english_word TEXT NOT NULL UNIQUE,
                    japanese_meaning TEXT NOT NULL,
                    created_at INTEGER
                )
            """)
            
            # 間違えた問題テーブル（新規追加）
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS wrong_questions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    question_type TEXT NOT NULL,  -- 'word_choice', 'word_listening', 'sentence'
                    question_content TEXT NOT NULL,  -- 英単語または英文
                    correct_answer TEXT NOT NULL,   -- 正解（日本語）
                    consecutive_correct INTEGER DEFAULT 0,  -- 連続正解数
                    created_at INTEGER,
                    UNIQUE(question_type, question_content)
                )
            """)
    
    def close(self):
        self.conn.close()
    
    # === 日付フィルタリング用ヘルパーメソッド（新規追加） ===
    def _get_date_filter_sql(self, start_date=None, end_date=None):
        """日付フィルタリング用のSQL条件文とパラメータを生成"""
        conditions = []
        params = []
        
        if start_date:
            start_timestamp = int(datetime.combine(start_date, datetime.min.time()).timestamp())
            conditions.append("created_at >= ?")
            params.append(start_timestamp)
        
        if end_date:
            end_timestamp = int(datetime.combine(end_date, datetime.max.time()).timestamp())
            conditions.append("created_at <= ?")
            params.append(end_timestamp)
        
        if conditions:
            return " AND " + " AND ".join(conditions), params
        return "", []
    
    # === 例文関連メソッド ===
    def add_sentence(self, english, japanese, ts):
        try:
            with self.conn:
                self.conn.execute("INSERT INTO sentences (english_sentence, japanese_translation, created_at) VALUES (?, ?, ?)", (english, japanese, ts))
            return True
        except sqlite3.IntegrityError:
            messagebox.showwarning("登録済み", "この英文は既に登録されています。")
            return False
    
    def get_all_sentences(self):
        return self.conn.execute("SELECT id, english_sentence, japanese_translation, created_at FROM sentences ORDER BY created_at DESC").fetchall()
    
    def delete_sentence(self, sentence_id):
        with self.conn:
            self.conn.execute("DELETE FROM sentences WHERE id = ?", (sentence_id,))
    
    def get_random_sentences(self, count, start_date=None, end_date=None):
        """日付フィルタ対応版"""
        date_filter, date_params = self._get_date_filter_sql(start_date, end_date)
        query = f"SELECT english_sentence, japanese_translation FROM sentences WHERE 1=1{date_filter} ORDER BY RANDOM() LIMIT ?"
        params = date_params + [count]
        return self.conn.execute(query, params).fetchall()
    
    def get_sentence_by_id(self, sentence_id):
        return self.conn.execute("SELECT english_sentence, japanese_translation FROM sentences WHERE id = ?", (sentence_id,)).fetchone()
    
    def update_sentence(self, sentence_id, english, japanese):
        """例文を更新"""
        with self.conn:
            self.conn.execute("UPDATE sentences SET english_sentence = ?, japanese_translation = ? WHERE id = ?", (english, japanese, sentence_id))

    def get_sentence_id_by_english(self, english_sentence):
        """英文から例文IDを取得"""
        result = self.conn.execute("SELECT id FROM sentences WHERE english_sentence = ?", (english_sentence,)).fetchone()
        return result[0] if result else None
    
    def get_sentences_count_by_date(self, start_date=None, end_date=None):
        """日付範囲内の例文数を取得"""
        date_filter, date_params = self._get_date_filter_sql(start_date, end_date)
        query = f"SELECT COUNT(*) FROM sentences WHERE 1=1{date_filter}"
        result = self.conn.execute(query, date_params).fetchone()
        return result[0] if result else 0
    
    # === 単語関連メソッド ===
    def add_word(self, english, japanese, ts):
        try:
            with self.conn:
                self.conn.execute("INSERT INTO words (english_word, japanese_meaning, created_at) VALUES (?, ?, ?)", (english, japanese, ts))
            return True
        except sqlite3.IntegrityError:
            messagebox.showwarning("登録済み", "この単語は既に登録されています。")
            return False
    
    def get_all_words(self):
        return self.conn.execute("SELECT id, english_word, japanese_meaning, created_at FROM words ORDER BY created_at DESC").fetchall()
    
    def delete_word(self, word_id):
        with self.conn:
            self.conn.execute("DELETE FROM words WHERE id = ?", (word_id,))
    
    def get_random_words(self, count, start_date=None, end_date=None):
        """日付フィルタ対応版"""
        date_filter, date_params = self._get_date_filter_sql(start_date, end_date)
        query = f"SELECT english_word, japanese_meaning FROM words WHERE 1=1{date_filter} ORDER BY RANDOM() LIMIT ?"
        params = date_params + [count]
        return self.conn.execute(query, params).fetchall()
    
    def get_word_by_id(self, word_id):
        return self.conn.execute("SELECT english_word, japanese_meaning FROM words WHERE id = ?", (word_id,)).fetchone()
    
    def update_word(self, word_id, english, japanese):
        """単語を更新"""
        with self.conn:
            self.conn.execute("UPDATE words SET english_word = ?, japanese_meaning = ? WHERE id = ?", (english, japanese, word_id))

    def get_word_id_by_english(self, english_word):
        """英単語から単語IDを取得"""
        result = self.conn.execute("SELECT id FROM words WHERE english_word = ?", (english_word,)).fetchone()
        return result[0] if result else None
    
    def get_words_count_by_date(self, start_date=None, end_date=None):
        """日付範囲内の単語数を取得"""
        date_filter, date_params = self._get_date_filter_sql(start_date, end_date)
        query = f"SELECT COUNT(*) FROM words WHERE 1=1{date_filter}"
        result = self.conn.execute(query, date_params).fetchone()
        return result[0] if result else 0
    
    def get_random_word_choices(self, exclude_meaning, n):
        """単語クイズ用の選択肢を取得"""
        base = exclude_meaning.split(';')[0].strip().lower()
        base_normalized = base.replace('・', '').replace(' ', '').replace('、', '').replace('。', '')
        
        rows = self.conn.execute(
            "SELECT japanese_meaning FROM words WHERE japanese_meaning != ? ORDER BY RANDOM() LIMIT 50", 
            (exclude_meaning,)
        ).fetchall()
        
        uniq, seen = [], {base_normalized}
        
        for (meaning,) in rows:
            if not meaning: continue
            
            cand = meaning.split(';')[0].strip()
            if not cand: continue
            
            norm = cand.lower().replace('・', '').replace(' ', '').replace('、', '').replace('。', '')
            
            if norm in seen or norm == base_normalized: continue
            if base_normalized in norm or norm in base_normalized: continue
            if len(norm) < 2: continue
            
            uniq.append(cand)
            seen.add(norm)
            
            if len(uniq) >= n: break
        
        # 十分な選択肢が取れない場合はダミーを追加
        while len(uniq) < n:
            dummy_options = ["該当なし", "不明", "その他", "関連語なし"]
            for dummy in dummy_options:
                if dummy not in uniq and len(uniq) < n:
                    uniq.append(dummy)
        
        return uniq[:n]

    # === 間違い問題管理メソッド（新規追加） ===
    def add_wrong_question(self, question_type, question_content, correct_answer):
        """間違えた問題を追加または更新"""
        try:
            with self.conn:
                self.conn.execute("""
                    INSERT OR REPLACE INTO wrong_questions 
                    (question_type, question_content, correct_answer, consecutive_correct, created_at) 
                    VALUES (?, ?, ?, 0, ?)
                """, (question_type, question_content, correct_answer, int(datetime.now().timestamp())))
        except Exception as e:
            print(f"Error adding wrong question: {e}")

    def update_wrong_question_score(self, question_type, question_content, is_correct):
        """間違えた問題の正解スコアを更新"""
        with self.conn:
            if is_correct:
                # 正解の場合、連続正解数をインクリメント
                self.conn.execute("""
                    UPDATE wrong_questions 
                    SET consecutive_correct = consecutive_correct + 1 
                    WHERE question_type = ? AND question_content = ?
                """, (question_type, question_content))
                
                # 3回連続正解したら削除
                self.conn.execute("""
                    DELETE FROM wrong_questions 
                    WHERE question_type = ? AND question_content = ? AND consecutive_correct >= 3
                """, (question_type, question_content))
            else:
                # 不正解の場合、連続正解数をリセット
                self.conn.execute("""
                    UPDATE wrong_questions 
                    SET consecutive_correct = 0 
                    WHERE question_type = ? AND question_content = ?
                """, (question_type, question_content))

    def get_wrong_questions(self, question_type, count=None):
        """間違えた問題を取得"""
        if count:
            query = """
                SELECT question_content, correct_answer, consecutive_correct 
                FROM wrong_questions 
                WHERE question_type = ? 
                ORDER BY RANDOM() 
                LIMIT ?
            """
            return self.conn.execute(query, (question_type, count)).fetchall()
        else:
            query = """
                SELECT question_content, correct_answer, consecutive_correct 
                FROM wrong_questions 
                WHERE question_type = ? 
                ORDER BY RANDOM()
            """
            return self.conn.execute(query, (question_type,)).fetchall()

    def get_wrong_questions_count(self, question_type):
        """間違えた問題の数を取得"""
        result = self.conn.execute(
            "SELECT COUNT(*) FROM wrong_questions WHERE question_type = ?", 
            (question_type,)
        ).fetchone()
        return result[0] if result else 0

# --- メインアプリケーションクラス ---
class App(tk.Tk):
    def __init__(self, config_data):
        super().__init__()
        self.config_data = config_data
        self.db_path = config_data.get("db_path", DEFAULT_DB_FILE)
        self.db = DatabaseManager(self.db_path)
        
        # DeepL翻訳設定
        self.translator = None
        # 編集中のアイテムID（新機能）
        self.editing_word_id = None
        self.editing_sentence_id = None
        
        # 日付フィルタ用変数（新規追加）
        self.filter_start_date = None
        self.filter_end_date = None
        
        self.title(APP_TITLE)
        self.geometry("1000x750")
        self.create_menu()
        self.create_widgets()
        self.refresh_all_lists()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_menu(self):
        """メニューバーを作成"""
        self.menu_bar = tk.Menu(self)
        self.configure(menu=self.menu_bar)
        
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="ファイル", menu=file_menu)
        
        file_menu.add_command(label="翻訳設定", command=self.show_deepl_settings)
        file_menu.add_separator()
        file_menu.add_command(label="音声ファイル作成", command=self.create_sound_files)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self.on_closing)

    def show_deepl_settings(self):
        """翻訳設定ダイアログ"""
        dialog = tk.Toplevel(self)
        dialog.title("翻訳設定")
        dialog.geometry("500x550")
        dialog.transient(self)
        dialog.grab_set()
        
        # APIキー入力
        ttk.Label(dialog, text="DeepL API キー:").pack(pady=10)
        api_key_var = tk.StringVar(value=self.config_data.get("deepl_api_key", ""))
        api_key_entry = ttk.Entry(dialog, textvariable=api_key_var, width=60)
        api_key_entry.pack(pady=5)
        
        # 説明文
        info_text = """APIキーの取得方法：
1. https://www.deepl.com/pro-api にアクセス
2. アカウントを作成（無料版もあります）
3. APIキーを取得してここに入力

公式DeepLライブラリを使用します。
・無料版：月間50万文字まで
・有料版：従量課金制

v3.6での改善点：
・Google翻訳ボタンを追加
・deep-translatorライブラリを使用
・無料で利用可能
・DeepLとGoogle翻訳を選択可能

必要ライブラリ：
pip install deepl edge-tts tkcalendar deep-translator

Google翻訳について：
・無料で利用可能
・APIキー不要
・ただし大量翻訳時は制限あり
・deep-translatorライブラリ経由で利用"""
        ttk.Label(dialog, text=info_text, justify=tk.LEFT, wraplength=450).pack(pady=10)
        
        # テスト機能
        def test_api():
            test_key = api_key_var.get().strip()
            if not test_key:
                messagebox.showerror("エラー", "APIキーを入力してください")
                return
            
            if not HAS_TRANSLATOR:
                messagebox.showerror("エラー", "DeepLライブラリがインストールされていません。")
                return
            
            # 長文でテスト
            test_text = "An Israeli strike in Gaza City late Sunday night killed seven people including five staff members from the news network Al Jazeera, in an attack condemned by press freedom advocates and the United Nations human rights office."
            
            try:
                test_translator = deepl.Translator(test_key)
                
                # 最新の修正方法でテスト
                if hasattr(deepl, 'SplitSentences'):
                    result = test_translator.translate_text(
                        test_text, 
                        target_lang="JA",
                        split_sentences=deepl.SplitSentences.OFF
                    )
                else:
                    result = test_translator.translate_text(
                        test_text, 
                        target_lang="JA",
                        split_sentences="off"
                    )
                
                messagebox.showinfo("テスト成功", f"DeepL APIキーは有効です！\n完全翻訳テスト結果:\n{result}")
            except deepl.AuthorizationException:
                messagebox.showerror("認証エラー", "APIキーが無効です。")
            except Exception as e:
                messagebox.showerror("テスト失敗", f"エラー: {e}")

        def test_google():
            """Google翻訳のテスト"""
            if not HAS_GOOGLE_TRANSLATOR:
                messagebox.showerror("エラー", "deep-translatorライブラリがインストールされていません。\npip install deep-translator を実行してください。")
                return
            
            test_text = "Hello world"
            try:
                translator = GoogleTranslator(source='en', target='ja')
                result = translator.translate(test_text)
                messagebox.showinfo("テスト成功", f"Google翻訳は利用可能です！\nテスト結果: {test_text} → {result}")
            except Exception as e:
                messagebox.showerror("テスト失敗", f"Google翻訳エラー: {e}")
        
        # ボタンフレーム
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="DeepL APIキーをテスト", command=test_api).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Google翻訳をテスト", command=test_google).pack(side=tk.LEFT, padx=5)
        
        def save_settings():
            self.config_data["deepl_api_key"] = api_key_var.get().strip()
            messagebox.showinfo("設定完了", "翻訳設定が保存されました。")
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save_settings).pack(side=tk.LEFT, padx=5)

    def create_sound_files(self):
        """音声ファイル作成機能"""
        if not HAS_TTS:
            messagebox.showerror("エラー", "edge-ttsライブラリがインストールされていません。\npip install edge-tts を実行してください。")
            return
        
        def create_files():
            try:
                asyncio.run(self._create_sound_files_async())
                messagebox.showinfo("作成完了", "音声ファイルが作成されました！\n- sound_correct.mp3 (正解音)\n- sound_incorrect.mp3 (不正解音)")
            except Exception as e:
                messagebox.showerror("作成失敗", f"音声ファイルの作成に失敗しました: {e}")
        
        threading.Thread(target=create_files, daemon=True).start()

    async def _create_sound_files_async(self):
        """音声ファイル作成の非同期処理"""
        # 正解音（ピンポン）
        correct_text = "ピンポン！正解です"
        communicate_correct = edge_tts.Communicate(correct_text, "ja-JP-NanamiNeural")
        await communicate_correct.save("sound_correct.mp3")
        
        # 不正解音（ブー）
        incorrect_text = "ブー！不正解です"
        communicate_incorrect = edge_tts.Communicate(incorrect_text, "ja-JP-KeitaNeural")
        await communicate_incorrect.save("sound_incorrect.mp3")

    def create_widgets(self):
        # メインコンテナ
        main_container = ttk.Frame(self, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # ノートブック（タブ）を作成
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 単語タブ
        self.word_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.word_frame, text="単語学習")
        self.create_word_widgets()
        
        # 例文タブ
        self.sentence_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sentence_frame, text="例文学習")
        self.create_sentence_widgets()
        
        # クイズタブ
        self.quiz_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.quiz_frame, text="クイズ")
        self.create_quiz_widgets()

    def create_word_widgets(self):
        """単語学習タブのウィジェット作成"""
        # 左右分割
        word_container = ttk.Frame(self.word_frame, padding=10)
        word_container.pack(fill=tk.BOTH, expand=True)
        
        # 左側フレーム（入力エリア）
        left_frame = ttk.Frame(word_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 編集状態表示
        self.word_edit_label = ttk.Label(left_frame, text="", foreground="blue", font=("", 10, "bold"))
        self.word_edit_label.pack(anchor='w', pady=(0, 5))
        
        # 英単語入力
        eng_frame = ttk.Frame(left_frame)
        eng_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(eng_frame, text="英単語:", font=("", 12)).pack(side=tk.LEFT)
        
        # 翻訳ボタンフレーム（★修正：複数のボタンを並べる）
        translate_btn_frame = ttk.Frame(eng_frame)
        translate_btn_frame.pack(side=tk.RIGHT)
        
        if HAS_GOOGLE_TRANSLATOR:
            ttk.Button(translate_btn_frame, text="Google翻訳", command=self.on_google_translate_word).pack(side=tk.RIGHT, padx=(0, 5))
        
        if HAS_TRANSLATOR:
            ttk.Button(translate_btn_frame, text="DeepL翻訳", command=self.on_deepl_translate_word).pack(side=tk.RIGHT)
        
        self.entry_word_english = ttk.Entry(left_frame, font=("", 14))
        self.entry_word_english.pack(fill=tk.X, pady=(0, 10))
        
        # 日本語意味入力
        ttk.Label(left_frame, text="日本語の意味:", font=("", 12)).pack(anchor='w', pady=(0, 5))
        self.entry_word_japanese = tk.Text(left_frame, height=4, wrap=tk.WORD, font=("", 11))
        self.entry_word_japanese.pack(fill=tk.X, pady=(0, 10))
        
        # ボタンエリア
        word_button_frame = ttk.Frame(left_frame)
        word_button_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(word_button_frame, text="単語を保存・更新", command=self.on_save_word).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, padx=(0, 5))
        ttk.Button(word_button_frame, text="クリア", command=self.on_clear_word_inputs).pack(side=tk.RIGHT, ipady=5)
        
        # 右側フレーム（リスト）
        right_frame = ttk.Frame(word_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        ttk.Label(right_frame, text="登録済み単語リスト（クリックで編集）").pack(anchor='w')
        
        # TreeView設定
        word_cols = ("ID", "英単語", "日本語")
        self.word_tree = ttk.Treeview(right_frame, columns=word_cols, show='headings', selectmode='browse')
        for col in word_cols:
            self.word_tree.heading(col, text=col)
        self.word_tree.column("ID", width=50, stretch=tk.NO)
        self.word_tree.column("英単語", width=150)
        self.word_tree.column("日本語", width=200)
        
        # TreeViewの選択イベントをバインド
        self.word_tree.bind("<<TreeviewSelect>>", self.on_word_select)
        
        self.word_tree.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # リスト操作ボタン
        word_list_btns_frame = ttk.Frame(right_frame)
        word_list_btns_frame.pack(fill=tk.X)
        ttk.Button(word_list_btns_frame, text="選択した単語を削除", command=self.on_delete_word).pack(side=tk.LEFT)

    def create_sentence_widgets(self):
        """例文学習タブのウィジェット作成"""
        # 左右分割
        sentence_container = ttk.Frame(self.sentence_frame, padding=10)
        sentence_container.pack(fill=tk.BOTH, expand=True)
        
        # 左側フレーム（入力エリア）
        left_frame = ttk.Frame(sentence_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 編集状態表示
        self.sentence_edit_label = ttk.Label(left_frame, text="", foreground="blue", font=("", 10, "bold"))
        self.sentence_edit_label.pack(anchor='w', pady=(0, 5))
        
        # 英語入力エリア
        eng_frame = ttk.Frame(left_frame)
        eng_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(eng_frame, text="英語の例文:", font=("", 12)).pack(side=tk.LEFT)
        
        # 翻訳ボタンフレーム（★修正：複数のボタンを並べる）
        translate_btn_frame = ttk.Frame(eng_frame)
        translate_btn_frame.pack(side=tk.RIGHT)
        
        if HAS_GOOGLE_TRANSLATOR:
            ttk.Button(translate_btn_frame, text="Google翻訳", command=self.on_google_translate_sentence).pack(side=tk.RIGHT, padx=(0, 5))
        
        if HAS_TRANSLATOR:
            ttk.Button(translate_btn_frame, text="DeepL翻訳", command=self.on_deepl_translate_sentence).pack(side=tk.RIGHT)

        self.entry_sentence_english = tk.Text(left_frame, height=5, wrap=tk.WORD, font=("", 11))
        self.entry_sentence_english.pack(fill=tk.X, pady=(0, 10))
        
        # 日本語入力エリア
        ttk.Label(left_frame, text="日本語訳:", font=("", 12)).pack(anchor='w', pady=(0, 5))
        self.entry_sentence_japanese = tk.Text(left_frame, height=5, wrap=tk.WORD, font=("", 11))
        self.entry_sentence_japanese.pack(fill=tk.X, pady=(0, 10))
        
        # ボタンエリア
        sentence_button_frame = ttk.Frame(left_frame)
        sentence_button_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(sentence_button_frame, text="例文を保存・更新", command=self.on_save_sentence).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, padx=(0, 5))
        ttk.Button(sentence_button_frame, text="クリア", command=self.on_clear_sentence_inputs).pack(side=tk.RIGHT, ipady=5)

        # 右側フレーム（リスト）
        right_frame = ttk.Frame(sentence_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        ttk.Label(right_frame, text="登録済み例文リスト（クリックで編集）").pack(anchor='w')
        
        # TreeView設定
        sentence_cols = ("ID", "英語の例文")
        self.sentence_tree = ttk.Treeview(right_frame, columns=sentence_cols, show='headings', selectmode='browse')
        for col in sentence_cols:
            self.sentence_tree.heading(col, text=col)
        self.sentence_tree.column("ID", width=50, stretch=tk.NO)
        self.sentence_tree.column("英語の例文", width=400)
        
        # TreeViewの選択イベントをバインド
        self.sentence_tree.bind("<<TreeviewSelect>>", self.on_sentence_select)
        
        self.sentence_tree.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # リスト操作ボタン
        sentence_list_btns_frame = ttk.Frame(right_frame)
        sentence_list_btns_frame.pack(fill=tk.X)
        ttk.Button(sentence_list_btns_frame, text="選択した例文を削除", command=self.on_delete_sentence).pack(side=tk.LEFT)

    def create_quiz_widgets(self):
        """クイズタブのウィジェット作成"""
        quiz_container = ttk.Frame(self.quiz_frame, padding=20)
        quiz_container.pack(fill=tk.BOTH, expand=True)
        
        # タイトル
        ttk.Label(quiz_container, text="クイズ選択", font=("", 16, "bold")).pack(pady=(0, 15))
        
        # === 日付フィルタセクション（新規追加） ===
        date_filter_frame = ttk.LabelFrame(quiz_container, text="日付フィルタ", padding=15)
        date_filter_frame.pack(fill=tk.X, pady=(0, 15))
        
        if HAS_CALENDAR:
            # 日付選択エリア
            date_select_frame = ttk.Frame(date_filter_frame)
            date_select_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Label(date_select_frame, text="開始日:").pack(side=tk.LEFT, padx=(0, 5))
            self.start_date_entry = DateEntry(date_select_frame, width=12, background='darkblue',
                                            foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
            self.start_date_entry.pack(side=tk.LEFT, padx=(0, 20))
            
            ttk.Label(date_select_frame, text="終了日:").pack(side=tk.LEFT, padx=(0, 5))
            self.end_date_entry = DateEntry(date_select_frame, width=12, background='darkblue',
                                          foreground='white', borderwidth=2, date_pattern='yyyy/mm/dd')
            self.end_date_entry.pack(side=tk.LEFT, padx=(0, 20))
            
            # プリセットボタン
            preset_frame = ttk.Frame(date_filter_frame)
            preset_frame.pack(fill=tk.X, pady=(0, 10))
            
            ttk.Button(preset_frame, text="全期間", command=self.set_all_period).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(preset_frame, text="最近1週間", command=self.set_last_week).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(preset_frame, text="最近1ヶ月", command=self.set_last_month).pack(side=tk.LEFT, padx=(0, 5))
            ttk.Button(preset_frame, text="今日", command=self.set_today).pack(side=tk.LEFT, padx=(0, 5))
            
            # 日付フィルタ統計
            self.date_stats_label = ttk.Label(date_filter_frame, text="", font=("", 10))
            self.date_stats_label.pack(anchor='w', pady=(5, 0))
            
            # デフォルトは全期間
            self.set_all_period()
        else:
            ttk.Label(date_filter_frame, text="日付フィルタ機能を使用するには tkcalendar をインストールしてください。\npip install tkcalendar", 
                     foreground="red", font=("", 10)).pack()
        
        # 単語クイズセクション（拡張版）
        word_quiz_frame = ttk.LabelFrame(quiz_container, text="単語クイズ", padding=15)
        word_quiz_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(word_quiz_frame, text="登録した単語からクイズを出題します").pack(anchor='w', pady=(0, 10))
        
        # 共通設定
        word_common_frame = ttk.Frame(word_quiz_frame)
        word_common_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(word_common_frame, text="問題数:").pack(side=tk.LEFT, padx=(0, 5))
        self.word_quiz_count = tk.IntVar(value=10)
        vcmd1 = (self.register(lambda P: str.isdigit(P) or P == ""), '%P')
        ttk.Entry(word_common_frame, textvariable=self.word_quiz_count, width=5, validate='key', validatecommand=vcmd1).pack(side=tk.LEFT, padx=(0, 20))
        
        # クイズ形式選択ボタン
        word_quiz_buttons = ttk.Frame(word_quiz_frame)
        word_quiz_buttons.pack(fill=tk.X)
        
        ttk.Button(word_quiz_buttons, text="4択クイズを開始", command=self.on_word_quiz_choice).pack(side=tk.LEFT, padx=(0, 5), ipady=5)
        ttk.Button(word_quiz_buttons, text="リスニングクイズを開始", command=self.on_word_quiz_listening).pack(side=tk.LEFT, padx=(0, 5), ipady=5)
        
        # 間違い復習ボタン（新規追加）
        wrong_word_frame = ttk.Frame(word_quiz_buttons)
        wrong_word_frame.pack(side=tk.LEFT, padx=(20, 0))
        
        self.wrong_word_choice_btn = ttk.Button(wrong_word_frame, text="間違い4択復習", command=self.on_wrong_word_quiz_choice)
        self.wrong_word_choice_btn.pack(side=tk.TOP, pady=(0, 2), ipady=3)
        
        self.wrong_word_listening_btn = ttk.Button(wrong_word_frame, text="間違いリスニング復習", command=self.on_wrong_word_quiz_listening)
        self.wrong_word_listening_btn.pack(side=tk.TOP, ipady=3)
        
        # 例文クイズセクション
        sentence_quiz_frame = ttk.LabelFrame(quiz_container, text="例文クイズ（リスニング形式）", padding=15)
        sentence_quiz_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(sentence_quiz_frame, text="登録した例文をリスニングクイズで学習します").pack(anchor='w', pady=(0, 10))
        
        sentence_quiz_settings = ttk.Frame(sentence_quiz_frame)
        sentence_quiz_settings.pack(fill=tk.X)
        
        ttk.Label(sentence_quiz_settings, text="問題数:").pack(side=tk.LEFT, padx=(0, 5))
        self.sentence_quiz_count = tk.IntVar(value=5)
        vcmd2 = (self.register(lambda P: str.isdigit(P) or P == ""), '%P')
        ttk.Entry(sentence_quiz_settings, textvariable=self.sentence_quiz_count, width=5, validate='key', validatecommand=vcmd2).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(sentence_quiz_settings, text="例文クイズを開始", command=self.on_sentence_quiz).pack(side=tk.LEFT, padx=(0, 20))
        
        # 間違い復習ボタン（例文用）
        self.wrong_sentence_btn = ttk.Button(sentence_quiz_settings, text="間違い例文復習", command=self.on_wrong_sentence_quiz)
        self.wrong_sentence_btn.pack(side=tk.LEFT)
        
        # 統計情報
        stats_frame = ttk.LabelFrame(quiz_container, text="学習統計", padding=15)
        stats_frame.pack(fill=tk.X)
        
        self.stats_label = ttk.Label(stats_frame, text="", font=("", 11))
        self.stats_label.pack(anchor='w')
        
        # 統計を更新
        self.update_stats()

    # === 日付フィルタ関連メソッド（新規追加） ===
    def set_all_period(self):
        """全期間を設定"""
        if not HAS_CALENDAR:
            return
        # データベースから最古と最新の日付を取得
        oldest_word = self.db.conn.execute("SELECT MIN(created_at) FROM words").fetchone()[0]
        oldest_sentence = self.db.conn.execute("SELECT MIN(created_at) FROM sentences").fetchone()[0]
        newest_word = self.db.conn.execute("SELECT MAX(created_at) FROM words").fetchone()[0]
        newest_sentence = self.db.conn.execute("SELECT MAX(created_at) FROM sentences").fetchone()[0]
        
        # 最古と最新を決定
        oldest = min(filter(None, [oldest_word, oldest_sentence])) if any([oldest_word, oldest_sentence]) else None
        newest = max(filter(None, [newest_word, newest_sentence])) if any([newest_word, newest_sentence]) else None
        
        if oldest:
            self.start_date_entry.set_date(datetime.fromtimestamp(oldest).date())
            self.filter_start_date = None  # 全期間の場合はNone
        else:
            self.start_date_entry.set_date(datetime.now().date() - timedelta(days=365))
            self.filter_start_date = None
            
        if newest:
            self.end_date_entry.set_date(datetime.fromtimestamp(newest).date())
            self.filter_end_date = None  # 全期間の場合はNone
        else:
            self.end_date_entry.set_date(datetime.now().date())
            self.filter_end_date = None
        
        self.update_date_stats()

    def set_last_week(self):
        """最近1週間を設定"""
        if not HAS_CALENDAR:
            return
        end_date = datetime.now().date()
        start_date = end_date - timedelta(days=7)
        
        self.start_date_entry.set_date(start_date)
        self.end_date_entry.set_date(end_date)
        self.filter_start_date = start_date
        self.filter_end_date = end_date
        self.update_date_stats()

    def set_last_month(self):
        """最近1ヶ月を設定"""
        if not HAS_CALENDAR:
            return
        end_date = datetime.now().date()
        start_date = end_date - timedelta(days=30)
        
        self.start_date_entry.set_date(start_date)
        self.end_date_entry.set_date(end_date)
        self.filter_start_date = start_date
        self.filter_end_date = end_date
        self.update_date_stats()

    def set_today(self):
        """今日を設定"""
        if not HAS_CALENDAR:
            return
        today = datetime.now().date()
        
        self.start_date_entry.set_date(today)
        self.end_date_entry.set_date(today)
        self.filter_start_date = today
        self.filter_end_date = today
        self.update_date_stats()

    def update_date_filter(self):
        """日付フィルタを更新"""
        if not HAS_CALENDAR:
            return
        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()
            
            if start_date > end_date:
                messagebox.showwarning("日付エラー", "開始日は終了日より前に設定してください。")
                return
            
            self.filter_start_date = start_date
            self.filter_end_date = end_date
            self.update_date_stats()
        except:
            pass

    def update_date_stats(self):
        """日付フィルタの統計を更新"""
        if not HAS_CALENDAR:
            return
        try:
            word_count = self.db.get_words_count_by_date(self.filter_start_date, self.filter_end_date)
            sentence_count = self.db.get_sentences_count_by_date(self.filter_start_date, self.filter_end_date)
            
            if self.filter_start_date and self.filter_end_date:
                if self.filter_start_date == self.filter_end_date:
                    period_text = f"{self.filter_start_date.strftime('%Y/%m/%d')}"
                else:
                    period_text = f"{self.filter_start_date.strftime('%Y/%m/%d')} ～ {self.filter_end_date.strftime('%Y/%m/%d')}"
            else:
                period_text = "全期間"
            
            stats_text = f"対象期間：{period_text}\n単語：{word_count}個　例文：{sentence_count}個"
            self.date_stats_label.config(text=stats_text)
        except:
            self.date_stats_label.config(text="日付統計取得エラー")

    def get_current_date_filter(self):
        """現在の日付フィルタを取得"""
        if not HAS_CALENDAR:
            return None, None
        
        try:
            self.update_date_filter()
            return self.filter_start_date, self.filter_end_date
        except:
            return None, None

    def update_stats(self):
        """学習統計を更新"""
        word_count = len(self.db.get_all_words())
        sentence_count = len(self.db.get_all_sentences())
        
        # 間違い問題の数も取得
        wrong_word_choice_count = self.db.get_wrong_questions_count('word_choice')
        wrong_word_listening_count = self.db.get_wrong_questions_count('word_listening')
        wrong_sentence_count = self.db.get_wrong_questions_count('sentence')
        
        stats_text = f"""全体統計:
・登録済み単語: {word_count}個
・登録済み例文: {sentence_count}個

復習待ち問題:
・単語4択: {wrong_word_choice_count}問
・単語リスニング: {wrong_word_listening_count}問
・例文リスニング: {wrong_sentence_count}問"""
        
        self.stats_label.config(text=stats_text)
        
        # ボタンの状態を更新
        if hasattr(self, 'wrong_word_choice_btn'):
            self.wrong_word_choice_btn.config(
                state="normal" if wrong_word_choice_count > 0 else "disabled",
                text=f"間違い4択復習 ({wrong_word_choice_count})"
            )
        if hasattr(self, 'wrong_word_listening_btn'):
            self.wrong_word_listening_btn.config(
                state="normal" if wrong_word_listening_count > 0 else "disabled",
                text=f"間違いリスニング復習 ({wrong_word_listening_count})"
            )
        if hasattr(self, 'wrong_sentence_btn'):
            self.wrong_sentence_btn.config(
                state="normal" if wrong_sentence_count > 0 else "disabled",
                text=f"間違い例文復習 ({wrong_sentence_count})"
            )

        # 日付統計も更新
        if HAS_CALENDAR:
            self.update_date_stats()

    # === 翻訳関連メソッド（★修正：DeepLとGoogle翻訳を分離） ===
    def translate_deepl(self, text):
        """DeepL翻訳を使用（カンマ問題修正版）"""
        if not text:
            return None
        
        if not HAS_TRANSLATOR:
            return "（DeepL翻訳ライブラリがインストールされていません。pip install deepl を実行してください）"
        
        deepl_api_key = self.config_data.get("deepl_api_key", "").strip()
        if not deepl_api_key:
            return "（DeepL APIキーが設定されていません。メニューから「翻訳設定」でAPIキーを設定してください）"
        
        try:
            translator = deepl.Translator(deepl_api_key)
            
            # テキストの前処理
            cleaned_text = text.strip()
            cleaned_text = re.sub(r'\r\n', '\n', cleaned_text)
            cleaned_text = re.sub(r'\n\s*\n', '\n', cleaned_text)
            cleaned_text = re.sub(r'[ \t]+', ' ', cleaned_text)
            
            # 方法1: enumを使用（最新版の場合）
            try:
                if hasattr(deepl, 'SplitSentences'):
                    result = translator.translate_text(
                        cleaned_text, 
                        target_lang="JA",
                        split_sentences=deepl.SplitSentences.OFF,
                        preserve_formatting=True
                    )
                    return str(result)
            except:
                pass
            
            # 方法2: 文字列指定
            try:
                result = translator.translate_text(
                    cleaned_text, 
                    target_lang="JA",
                    split_sentences="off",
                    preserve_formatting=True
                )
                return str(result)
            except:
                pass
            
            # 方法3: 数値指定
            result = translator.translate_text(
                cleaned_text, 
                target_lang="JA",
                split_sentences=0
            )
            return str(result)
                
        except deepl.QuotaExceededException:
            return "（DeepL APIの月間制限に達しました）"
        except deepl.AuthorizationException:
            return "（DeepL APIキーが無効です。正しいAPIキーを設定してください）"
        except Exception as e:
            return f"（DeepL翻訳に失敗しました: {e}）"

    def translate_google(self, text):
        """Google翻訳を使用（deep-translator）"""
        if not text:
            return None
        
        if not HAS_GOOGLE_TRANSLATOR:
            return "（deep-translatorライブラリがインストールされていません。pip install deep-translator を実行してください）"
        
        try:
            # テキストの前処理
            cleaned_text = text.strip()
            cleaned_text = re.sub(r'\r\n', '\n', cleaned_text)
            cleaned_text = re.sub(r'\n\s*\n', '\n', cleaned_text)
            cleaned_text = re.sub(r'[ \t]+', ' ', cleaned_text)
            
            # Google翻訳を実行
            translator = GoogleTranslator(source='en', target='ja')
            result = translator.translate(cleaned_text)
            return result
                
        except Exception as e:
            return f"（Google翻訳に失敗しました: {e}）"

    # === 単語関連メソッド ===
    def on_deepl_translate_word(self):
        """単語DeepL翻訳ボタンが押されたときの処理"""
        word_text = self.entry_word_english.get().strip()
        if not word_text:
            messagebox.showwarning("入力エラー", "翻訳する単語を入力してください。")
            return
        
        try:
            translated_text = self.translate_deepl(word_text)
            if translated_text and translated_text.startswith("（") and translated_text.endswith("）"):
                messagebox.showerror("翻訳エラー", translated_text)
            else:
                self.entry_word_japanese.delete("1.0", tk.END)
                self.entry_word_japanese.insert("1.0", translated_text)
        except Exception as e:
            messagebox.showerror("翻訳エラー", f"翻訳中にエラーが発生しました。\n{e}")

    def on_google_translate_word(self):
        """単語Google翻訳ボタンが押されたときの処理"""
        word_text = self.entry_word_english.get().strip()
        if not word_text:
            messagebox.showwarning("入力エラー", "翻訳する単語を入力してください。")
            return
        
        try:
            translated_text = self.translate_google(word_text)
            if translated_text and translated_text.startswith("（") and translated_text.endswith("）"):
                messagebox.showerror("翻訳エラー", translated_text)
            else:
                self.entry_word_japanese.delete("1.0", tk.END)
                self.entry_word_japanese.insert("1.0", translated_text)
        except Exception as e:
            messagebox.showerror("翻訳エラー", f"翻訳中にエラーが発生しました。\n{e}")

    def on_save_word(self):
        """単語保存・更新"""
        english = self.entry_word_english.get().strip()
        japanese = self.entry_word_japanese.get("1.0", tk.END).strip()
        
        if not english or not japanese:
            messagebox.showwarning("入力エラー", "英単語と日本語の意味の両方を入力してください。")
            return
        
        if self.editing_word_id:
            # 更新処理
            self.db.update_word(self.editing_word_id, english, japanese)
            operation = "更新"
            self.editing_word_id = None
            self.word_edit_label.config(text="")
        else:
            # 既存チェック
            existing_id = self.db.get_word_id_by_english(english)
            if existing_id:
                # 既存単語の更新
                self.db.update_word(existing_id, english, japanese)
                operation = "更新"
            else:
                # 新規追加
                if not self.db.add_word(english, japanese, int(datetime.now().timestamp())):
                    return
                operation = "保存"
        
        # 入力フィールドをクリア
        self.entry_word_english.delete(0, tk.END)
        self.entry_word_japanese.delete("1.0", tk.END)
        
        # リストを更新
        self.refresh_word_list()
        self.update_stats()
        
        # ポップアップは表示しない
        print(f"単語が{operation}されました: {english}")

    def on_clear_word_inputs(self):
        """単語入力フィールドをクリア"""
        self.entry_word_english.delete(0, tk.END)
        self.entry_word_japanese.delete("1.0", tk.END)
        self.editing_word_id = None
        self.word_edit_label.config(text="")

    def on_word_select(self, event):
        """単語が選択されたときの処理（編集モード）"""
        selected_item = self.word_tree.focus()
        if not selected_item:
            return
        
        item_values = self.word_tree.item(selected_item, 'values')
        if not item_values:
            return
        
        word_id = item_values[0]
        word_data = self.db.get_word_by_id(word_id)
        if word_data:
            english_word, japanese_meaning = word_data
            
            self.entry_word_english.delete(0, tk.END)
            self.entry_word_english.insert(0, english_word)
            
            self.entry_word_japanese.delete("1.0", tk.END)
            self.entry_word_japanese.insert("1.0", japanese_meaning)
            
            # 編集モード設定
            self.editing_word_id = word_id
            self.word_edit_label.config(text=f"編集中: {english_word}")

    def on_delete_word(self):
        """単語削除"""
        selected_item = self.word_tree.focus()
        if not selected_item:
            messagebox.showwarning("選択なし", "削除する単語をリストから選択してください。")
            return
        
        item_values = self.word_tree.item(selected_item, 'values')
        if messagebox.askyesno("削除の確認", f"以下の単語を削除しますか？\n\n{item_values[1]} → {item_values[2]}"):
            self.db.delete_word(item_values[0])
            self.refresh_word_list()
            self.on_clear_word_inputs()
            self.update_stats()

    def refresh_word_list(self):
        """単語リスト更新"""
        for item in self.word_tree.get_children():
            self.word_tree.delete(item)
        
        for row in self.db.get_all_words():
            # row = (id, english_word, japanese_meaning, created_at)
            self.word_tree.insert("", tk.END, values=(row[0], row[1], row[2][:30]))

    # === 例文関連メソッド ===
    def on_deepl_translate_sentence(self):
        """例文DeepL翻訳ボタンが押されたときの処理"""
        sentence_text = self.entry_sentence_english.get("1.0", tk.END).strip()
        if not sentence_text:
            messagebox.showwarning("入力エラー", "翻訳する英文を入力してください。")
            return
        
        try:
            translated_text = self.translate_deepl(sentence_text)
            if translated_text and translated_text.startswith("（") and translated_text.endswith("）"):
                messagebox.showerror("翻訳エラー", translated_text)
            else:
                self.entry_sentence_japanese.delete("1.0", tk.END)
                self.entry_sentence_japanese.insert("1.0", translated_text)
        except Exception as e:
            messagebox.showerror("翻訳エラー", f"翻訳中にエラーが発生しました。\n{e}")

    def on_google_translate_sentence(self):
        """例文Google翻訳ボタンが押されたときの処理"""
        sentence_text = self.entry_sentence_english.get("1.0", tk.END).strip()
        if not sentence_text:
            messagebox.showwarning("入力エラー", "翻訳する英文を入力してください。")
            return
        
        try:
            translated_text = self.translate_google(sentence_text)
            if translated_text and translated_text.startswith("（") and translated_text.endswith("）"):
                messagebox.showerror("翻訳エラー", translated_text)
            else:
                self.entry_sentence_japanese.delete("1.0", tk.END)
                self.entry_sentence_japanese.insert("1.0", translated_text)
        except Exception as e:
            messagebox.showerror("翻訳エラー", f"翻訳中にエラーが発生しました。\n{e}")

    def on_save_sentence(self):
        """例文保存・更新"""
        english = self.entry_sentence_english.get("1.0", tk.END).strip()
        japanese = self.entry_sentence_japanese.get("1.0", tk.END).strip()
        
        if not english or not japanese:
            messagebox.showwarning("入力エラー", "英語の例文と日本語訳の両方を入力してください。")
            return
        
        if self.editing_sentence_id:
            # 更新処理
            self.db.update_sentence(self.editing_sentence_id, english, japanese)
            operation = "更新"
            self.editing_sentence_id = None
            self.sentence_edit_label.config(text="")
        else:
            # 既存チェック
            existing_id = self.db.get_sentence_id_by_english(english)
            if existing_id:
                # 既存例文の更新
                self.db.update_sentence(existing_id, english, japanese)
                operation = "更新"
            else:
                # 新規追加
                if not self.db.add_sentence(english, japanese, int(datetime.now().timestamp())):
                    return
                operation = "保存"
        
        # 入力フィールドをクリア
        self.entry_sentence_english.delete("1.0", tk.END)
        self.entry_sentence_japanese.delete("1.0", tk.END)
        
        # リストを更新
        self.refresh_sentence_list()
        self.update_stats()
        
        # ポップアップは表示しない
        print(f"例文が{operation}されました: {english[:30]}...")

    def on_clear_sentence_inputs(self):
        """例文入力フィールドをクリア"""
        self.entry_sentence_english.delete("1.0", tk.END)
        self.entry_sentence_japanese.delete("1.0", tk.END)
        self.editing_sentence_id = None
        self.sentence_edit_label.config(text="")

    def on_sentence_select(self, event):
        """例文が選択されたときの処理（編集モード）"""
        selected_item = self.sentence_tree.focus()
        if not selected_item:
            return
        
        item_values = self.sentence_tree.item(selected_item, 'values')
        if not item_values:
            return
        
        sentence_id = item_values[0]
        sentence_data = self.db.get_sentence_by_id(sentence_id)
        if sentence_data:
            english_sentence, japanese_translation = sentence_data
            
            self.entry_sentence_english.delete("1.0", tk.END)
            self.entry_sentence_english.insert("1.0", english_sentence)
            
            self.entry_sentence_japanese.delete("1.0", tk.END)
            self.entry_sentence_japanese.insert("1.0", japanese_translation)
            
            # 編集モード設定
            self.editing_sentence_id = sentence_id
            self.sentence_edit_label.config(text=f"編集中: {english_sentence[:30]}...")

    def on_delete_sentence(self):
        """例文削除"""
        selected_item = self.sentence_tree.focus()
        if not selected_item:
            messagebox.showwarning("選択なし", "削除する例文をリストから選択してください。")
            return
        
        item_values = self.sentence_tree.item(selected_item, 'values')
        if messagebox.askyesno("削除の確認", f"以下の例文を削除しますか？\n\n{item_values[1]}"):
            self.db.delete_sentence(item_values[0])
            self.refresh_sentence_list()
            self.on_clear_sentence_inputs()
            self.update_stats()

    def refresh_sentence_list(self):
        """例文リスト更新"""
        for item in self.sentence_tree.get_children():
            self.sentence_tree.delete(item)
        
        for row in self.db.get_all_sentences():
            self.sentence_tree.insert("", tk.END, values=(row[0], row[1]))

    def refresh_all_lists(self):
        """全リスト更新"""
        self.refresh_word_list()
        self.refresh_sentence_list()
        self.update_stats()

    # === クイズ関連メソッド（日付フィルタ対応版） ===
    def on_word_quiz_choice(self):
        """単語クイズ（4択形式）開始"""
        try:
            count = self.word_quiz_count.get()
            if count <= 0:
                messagebox.showerror("入力エラー", "問題数は1以上で指定してください。")
                return
        except tk.TclError:
            messagebox.showerror("入力エラー", "問題数が正しく入力されていません。")
            return
        
        # 日付フィルタを適用
        start_date, end_date = self.get_current_date_filter()
        questions = self.db.get_random_words(count, start_date, end_date)
        
        if not questions:
            messagebox.showinfo("クイズ中止", "指定した期間にクイズに出題できる単語がありません。")
            return
        
        if len(questions) < 4:
            messagebox.showinfo("クイズ中止", "4択クイズには最低4つの単語が必要です。")
            return
        
        WordQuizChoiceWindow(self, questions)

    def on_word_quiz_listening(self):
        """単語クイズ（リスニング形式）開始"""
        try:
            count = self.word_quiz_count.get()
            if count <= 0:
                messagebox.showerror("入力エラー", "問題数は1以上で指定してください。")
                return
        except tk.TclError:
            messagebox.showerror("入力エラー", "問題数が正しく入力されていません。")
            return
        
        # 日付フィルタを適用
        start_date, end_date = self.get_current_date_filter()
        questions = self.db.get_random_words(count, start_date, end_date)
        
        if not questions:
            messagebox.showinfo("クイズ中止", "指定した期間にクイズに出題できる単語がありません。")
            return
        
        WordQuizListeningWindow(self, questions)

    def on_sentence_quiz(self):
        """例文クイズ開始"""
        try:
            count = self.sentence_quiz_count.get()
            if count <= 0:
                messagebox.showerror("入力エラー", "問題数は1以上で指定してください。")
                return
        except tk.TclError:
            messagebox.showerror("入力エラー", "問題数が正しく入力されていません。")
            return
        
        # 日付フィルタを適用
        start_date, end_date = self.get_current_date_filter()
        questions = self.db.get_random_sentences(count, start_date, end_date)
        
        if not questions:
            messagebox.showinfo("クイズ中止", "指定した期間にクイズに出題できる例文がありません。")
            return
        
        SentenceQuizWindow(self, questions)

    # === 間違い復習クイズメソッド（新規追加） ===
    def on_wrong_word_quiz_choice(self):
        """間違えた単語クイズ（4択形式）開始"""
        questions = self.db.get_wrong_questions('word_choice')
        if not questions:
            messagebox.showinfo("復習なし", "復習対象の単語4択問題がありません。")
            return
        
        WrongWordQuizChoiceWindow(self, questions)

    def on_wrong_word_quiz_listening(self):
        """間違えた単語クイズ（リスニング形式）開始"""
        questions = self.db.get_wrong_questions('word_listening')
        if not questions:
            messagebox.showinfo("復習なし", "復習対象の単語リスニング問題がありません。")
            return
        
        WrongWordQuizListeningWindow(self, questions)

    def on_wrong_sentence_quiz(self):
        """間違えた例文クイズ開始"""
        questions = self.db.get_wrong_questions('sentence')
        if not questions:
            messagebox.showinfo("復習なし", "復習対象の例文問題がありません。")
            return
        
        WrongSentenceQuizWindow(self, questions)

    def on_closing(self):
        save_config(self.config_data)
        self.db.close()
        self.destroy()

# === 通常クイズウィンドウクラス（間違い記録機能付き） ===

# --- 単語クイズウィンドウクラス（4択形式｜間違い記録対応版） ---
class WordQuizChoiceWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__(master)
        self.master_app = master
        self.questions = questions
        self.total_questions = len(questions)
        self.current_q_index = 0
        self.score = 0
        self.is_speaking = False

        # ★追加：間違い復習用リストとボタン参照
        self.wrongs = []           # [(english_word, correct_meaning), ...]
        self.review_btn = None

        self.title("単語クイズ（4択形式）")
        self.geometry("600x520")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_var = tk.StringVar()
        ttk.Label(self, textvariable=self.progress_var, font=("", 12)).pack(pady=10)

        word_frame = ttk.Frame(self)
        word_frame.pack(pady=15)
        self.word_label = ttk.Label(word_frame, text="", font=("", 24, "bold"))
        self.word_label.pack(side=tk.LEFT, padx=(0, 10))

        if HAS_TTS:
            self.speak_button = ttk.Button(word_frame, text="🔊 発音", command=self.speak_current_word, width=8)
            self.speak_button.pack(side=tk.LEFT)

        self.feedback_var = tk.StringVar()
        ttk.Label(self, textvariable=self.feedback_var, font=("", 12)).pack(pady=10)

        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(pady=20, expand=True, fill=tk.BOTH)

        self.next_btn = ttk.Button(self, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(pady=10)

    def _play_sound(self, sound_path):
        """音声ファイル再生（非ブロッキング）"""
        try:
            winmm = ctypes.windll.winmm
            alias = f"sound_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(1500, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    def _speak_task(self, text_to_say):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text_to_say, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                self._play_sound(AUDIO_FILE_PATH)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError as e:
                        print(f"Error removing audio file: {e}")

        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_word(self):
        if HAS_TTS and not self.is_speaking:
            word_to_say = self.word_label.cget("text")
            threading.Thread(target=self._speak_task, args=(word_to_say,), daemon=True).start()

    def show_next_question(self):
        if self.current_q_index >= self.total_questions:
            self.show_final_score()
            return

        self.feedback_var.set("")
        self.next_btn.config(state="disabled")
        for btn in self.btn_frame.winfo_children():
            btn.destroy()

        q = self.questions[self.current_q_index]
        self.correct_answer = q[1]  # 日本語の意味
        word_to_show = q[0]         # 英単語

        self.progress_var.set(f"第 {self.current_q_index + 1} 問 / 全 {self.total_questions} 問 (正解: {self.score})")
        self.word_label.config(text=word_to_show)
        self.speak_current_word()

        # 選択肢生成
        choices = self.master_app.db.get_random_word_choices(self.correct_answer, 3)
        all_choices = [self.correct_answer] + choices

        # 重複除去
        unique_choices = []
        seen = set()
        for choice in all_choices:
            normalized = choice.lower().replace('・', '').replace(' ', '').replace('、', '').replace('。', '')
            if normalized not in seen and choice.strip():
                unique_choices.append(choice)
                seen.add(normalized)

        # 4つ確保
        while len(unique_choices) < 4:
            dummy_options = ["該当なし", "不明", "その他", "関連語なし"]
            for dummy in dummy_options:
                if dummy not in unique_choices:
                    unique_choices.append(dummy)
                    break

        unique_choices = unique_choices[:4]
        random.shuffle(unique_choices)

        # ボタン作成
        for i, choice in enumerate(unique_choices):
            btn = ttk.Button(self.btn_frame, text=f"{chr(65+i)}. {choice}",
                             command=lambda c=choice: self.check_answer(c))
            btn.pack(fill=tk.X, pady=8, padx=50, ipady=10)

        self.current_q_index += 1

    def check_answer(self, choice):
        is_correct = choice == self.correct_answer

        if is_correct:
            self.score += 1
            self.feedback_var.set("✅ 正解！")
            sound_file = SOUND_FILES["correct"]
        else:
            self.feedback_var.set(f"❌ 不正解。正解は「{self.correct_answer}」")
            sound_file = SOUND_FILES["incorrect"]
            # ★追加：間違いを記録（表示中の英単語と正解の意味）
            self.wrongs.append((self.word_label.cget("text"), self.correct_answer))
            # データベースにも記録
            self.master_app.db.add_wrong_question('word_choice', self.word_label.cget("text"), self.correct_answer)

        # 効果音再生
        if os.path.exists(sound_file):
            threading.Thread(target=self._play_sound, args=(sound_file,), daemon=True).start()

        for btn in self.btn_frame.winfo_children():
            btn.config(state="disabled")
        self.next_btn.config(state="normal")

        if HAS_TTS and hasattr(self, 'speak_button'):
            self.speak_button.config(state="normal")

        self.progress_var.set(f"第 {self.current_q_index} 問 / 全 {self.total_questions} 問 (正解: {self.score})")

    def show_final_score(self):
        self.word_label.config(text="クイズ終了！")
        self.feedback_var.set(f"最終結果: {self.total_questions} 問中 {self.score} 問正解！")
        for btn in self.btn_frame.winfo_children():
            btn.destroy()

        # ★追加：復習ボタン（間違いがある場合のみ）
        if self.wrongs and not self.review_btn:
            self.review_btn = ttk.Button(self, text=f"間違えた問題を復習（{len(self.wrongs)}件）", command=self.open_review)
            self.review_btn.pack(pady=(0, 6))

        self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")
        if HAS_TTS and hasattr(self, 'speak_button'):
            self.speak_button.config(state="disabled")

        # 統計更新
        self.master_app.update_stats()

    def open_review(self):
        # 既存の WordReviewWindow（単語復習用）を再利用
        WordReviewWindow(self, self.wrongs)

    def on_closing(self):
        self.destroy()


# --- 単語クイズウィンドウクラス（リスニング形式｜間違い記録対応） ---
class WordQuizListeningWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__()
        self.master_app = master
        self.questions = questions
        self.total = len(questions)
        self.current_idx = 0
        self.is_speaking = False
        self.score = 0            # 正答数
        self.wrongs = []          # ★追加：間違えた問題の [(word, meaning), ...]
        self.review_btn = None    # 終了時の復習ボタン参照

        self.title("単語リスニングクイズ")
        self.geometry("740x460")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_label = ttk.Label(self, text="", font=("", 12))
        self.progress_label.pack(pady=10)

        self.word_label = ttk.Label(self, text="...", font=("", 24, "bold"), justify=tk.CENTER)
        self.word_label.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.meaning_label = ttk.Label(self, text="（答えは下に表示されます）", font=("", 14), justify=tk.CENTER)
        self.meaning_label.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        self.judge_feedback = ttk.Label(self, text="", font=("", 11), foreground="blue")
        self.judge_feedback.pack(pady=(0, 6))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn_frame, text="🔊 もう一度聞く", command=self.speak_current_word)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.answer_btn = ttk.Button(btn_frame, text="答えを見る", command=self.show_answer)
        self.answer_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.correct_btn = ttk.Button(btn_frame, text="〇 正解", command=lambda: self.judge(True), state="disabled")
        self.correct_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.wrong_btn = ttk.Button(btn_frame, text="× 不正解", command=lambda: self.judge(False), state="disabled")
        self.wrong_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn_frame, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_next_question(self):
        if self.current_idx >= self.total:
            self.word_label.config(text="クイズ終了！")
            self.meaning_label.config(text=f"全{self.total}問お疲れ様でした。最終結果：{self.total} 問中 {self.score} 問正解")
            self.answer_btn.config(state="disabled")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.correct_btn.config(state="disabled")
            self.wrong_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")

            # ★追加：復習ボタン（間違いがあるときのみ）
            if self.wrongs and not self.review_btn:
                self.review_btn = ttk.Button(self, text=f"間違えた問題を復習（{len(self.wrongs)}件）", command=self.open_review)
                self.review_btn.pack(pady=(5, 12))

            # 統計更新
            self.master_app.update_stats()
            return

        q = self.questions[self.current_idx]
        self.current_word = q[0]
        self.current_meaning = q[1]

        self.progress_label.config(text=f"第 {self.current_idx + 1} 問 / 全 {self.total} 問（正解: {self.score}）")
        self.word_label.config(text=self.current_word)
        self.meaning_label.config(text="（下に日本語の意味が表示されます）")
        self.judge_feedback.config(text="")

        self.answer_btn.config(state="normal")
        self.next_btn.config(state="disabled")
        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        if hasattr(self, 'speak_btn'):
            self.speak_btn.config(state="normal")

        if HAS_TTS:
            self.speak_current_word()

        self.current_idx += 1

    def show_answer(self):
        self.meaning_label.config(text=self.current_meaning)
        self.answer_btn.config(state="disabled")
        self.correct_btn.config(state="normal")
        self.wrong_btn.config(state="normal")
        self.judge_feedback.config(text="正解なら『〇 正解』、間違いなら『× 不正解』を押してください。")

    def judge(self, is_correct: bool):
        if is_correct:
            self.score += 1
            self.judge_feedback.config(text="✅ 正解として記録しました")
            sound_file = SOUND_FILES["correct"]
        else:
            # ★追加：間違いを記録
            self.wrongs.append((self.current_word, self.current_meaning))
            self.master_app.db.add_wrong_question('word_listening', self.current_word, self.current_meaning)
            self.judge_feedback.config(text="❌ 不正解として記録しました")
            sound_file = SOUND_FILES["incorrect"]

        if os.path.exists(sound_file):
            self._play_sound(sound_file)

        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        self.next_btn.config(state="normal")
        self.progress_label.config(text=f"第 {self.current_idx} 問 / 全 {self.total} 問（正解: {self.score}）")

    def open_review(self):
        # 復習ウィンドウを開く（間違えた単語のみ）
        WordReviewWindow(self, self.wrongs)

    def _play_sound(self, sound_path, duration_ms=1500):
        try:
            winmm = ctypes.windll.winmm
            alias = f"se_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(duration_ms, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    # --- 既存のTTS処理 ---
    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"speech_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_word(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_word,), daemon=True).start()

    def on_closing(self):
        self.destroy()


# --- 例文クイズウィンドウクラス（リスニング形式｜間違い記録対応） ---
class SentenceQuizWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__()
        self.master_app = master
        self.questions = questions
        self.total = len(questions)
        self.current_idx = 0
        self.is_speaking = False
        self.score = 0
        self.wrongs = []          # ★追加：間違いリスト
        self.review_btn = None

        self.title("例文リスニングクイズ")
        self.geometry("760x480")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_label = ttk.Label(self, text="", font=("", 12))
        self.progress_label.pack(pady=10)

        self.english_label = ttk.Label(self, text="...", font=("", 16, "bold"), wraplength=700, justify=tk.CENTER)
        self.english_label.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.japanese_label = ttk.Label(self, text="（答えは下に表示されます）", font=("", 12), wraplength=700, justify=tk.CENTER)
        self.japanese_label.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        self.judge_feedback = ttk.Label(self, text="", font=("", 11), foreground="blue")
        self.judge_feedback.pack(pady=(0, 6))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn_frame, text="🔊 もう一度聞く", command=self.speak_current_sentence)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.answer_btn = ttk.Button(btn_frame, text="答えを見る", command=self.show_answer)
        self.answer_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.correct_btn = ttk.Button(btn_frame, text="〇 正解", command=lambda: self.judge(True), state="disabled")
        self.correct_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.wrong_btn = ttk.Button(btn_frame, text="× 不正解", command=lambda: self.judge(False), state="disabled")
        self.wrong_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn_frame, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_next_question(self):
        if self.current_idx >= self.total:
            self.english_label.config(text="クイズ終了！")
            self.japanese_label.config(text=f"全{self.total}問お疲れ様でした。最終結果：{self.total} 問中 {self.score} 問正解")
            self.answer_btn.config(state="disabled")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.correct_btn.config(state="disabled")
            self.wrong_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")

            # ★追加：復習ボタン（間違いがあるときのみ）
            if self.wrongs and not self.review_btn:
                self.review_btn = ttk.Button(self, text=f"間違えた問題を復習（{len(self.wrongs)}件）", command=self.open_review)
                self.review_btn.pack(pady=(5, 12))

            # 統計更新
            self.master_app.update_stats()
            return

        q = self.questions[self.current_idx]
        self.current_english = q[0]
        self.current_japanese = q[1]

        self.progress_label.config(text=f"第 {self.current_idx + 1} 問 / 全 {self.total} 問（正解: {self.score}）")
        self.english_label.config(text=self.current_english)
        self.japanese_label.config(text="（下に日本語訳が表示されます）")
        self.judge_feedback.config(text="")

        self.answer_btn.config(state="normal")
        self.next_btn.config(state="disabled")
        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        if hasattr(self, 'speak_btn'):
            self.speak_btn.config(state="normal")

        if HAS_TTS:
            self.speak_current_sentence()

        self.current_idx += 1

    def show_answer(self):
        self.japanese_label.config(text=self.current_japanese)
        self.answer_btn.config(state="disabled")
        self.correct_btn.config(state="normal")
        self.wrong_btn.config(state="normal")
        self.judge_feedback.config(text="正解なら『〇 正解』、間違いなら『× 不正解』を押してください。")

    def judge(self, is_correct: bool):
        if is_correct:
            self.score += 1
            self.judge_feedback.config(text="✅ 正解として記録しました")
            sound_file = SOUND_FILES["correct"]
        else:
            # ★追加：間違いを記録
            self.wrongs.append((self.current_english, self.current_japanese))
            self.master_app.db.add_wrong_question('sentence', self.current_english, self.current_japanese)
            self.judge_feedback.config(text="❌ 不正解として記録しました")
            sound_file = SOUND_FILES["incorrect"]

        if os.path.exists(sound_file):
            self._play_sound(sound_file)

        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        self.next_btn.config(state="normal")
        self.progress_label.config(text=f"第 {self.current_idx} 問 / 全 {self.total} 問（正解: {self.score}）")

    def open_review(self):
        # 復習ウィンドウを開く（間違えた例文のみ）
        SentenceReviewWindow(self, self.wrongs)

    def _play_sound(self, sound_path, duration_ms=1500):
        try:
            winmm = ctypes.windll.winmm
            alias = f"se_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(duration_ms, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    # --- 既存のTTS処理 ---
    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"speech_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_sentence(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_english,), daemon=True).start()

    def on_closing(self):
        self.destroy()

# === 間違い復習専用クイズウィンドウクラス（新規追加） ===

# --- 間違い単語クイズ（4択形式）---
class WrongWordQuizChoiceWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__(master)
        self.master_app = master
        self.questions = list(questions)  # [(question_content, correct_answer, consecutive_correct), ...]
        self.total_questions = len(questions)
        self.current_q_index = 0
        self.score = 0
        self.is_speaking = False
        self.used_questions = set()  # 既に出題した問題を記録

        self.title("間違い復習：単語4択クイズ")
        self.geometry("600x520")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_var = tk.StringVar()
        ttk.Label(self, textvariable=self.progress_var, font=("", 12)).pack(pady=10)

        word_frame = ttk.Frame(self)
        word_frame.pack(pady=15)
        self.word_label = ttk.Label(word_frame, text="", font=("", 24, "bold"))
        self.word_label.pack(side=tk.LEFT, padx=(0, 10))

        if HAS_TTS:
            self.speak_button = ttk.Button(word_frame, text="🔊 発音", command=self.speak_current_word, width=8)
            self.speak_button.pack(side=tk.LEFT)

        self.feedback_var = tk.StringVar()
        ttk.Label(self, textvariable=self.feedback_var, font=("", 12)).pack(pady=10)

        self.btn_frame = ttk.Frame(self)
        self.btn_frame.pack(pady=20, expand=True, fill=tk.BOTH)

        self.next_btn = ttk.Button(self, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(pady=10)

    def _play_sound(self, sound_path):
        """音声ファイル再生（非ブロッキング）"""
        try:
            winmm = ctypes.windll.winmm
            alias = f"sound_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(1500, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    def _speak_task(self, text_to_say):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text_to_say, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                self._play_sound(AUDIO_FILE_PATH)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError as e:
                        print(f"Error removing audio file: {e}")

        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_word(self):
        if HAS_TTS and not self.is_speaking:
            word_to_say = self.word_label.cget("text")
            threading.Thread(target=self._speak_task, args=(word_to_say,), daemon=True).start()

    def show_next_question(self):
        # 利用可能な問題をフィルタリング（まだ出題していない問題）
        available_questions = [q for i, q in enumerate(self.questions) if i not in self.used_questions]
        
        if not available_questions:
            self.show_final_score()
            return

        self.feedback_var.set("")
        self.next_btn.config(state="disabled")
        for btn in self.btn_frame.winfo_children():
            btn.destroy()

        # ランダムに1つ選択
        q = random.choice(available_questions)
        # 使用済みに追加
        original_index = self.questions.index(q)
        self.used_questions.add(original_index)
        
        self.current_question = q[0]  # 英単語
        self.correct_answer = q[1]   # 正解
        self.consecutive_correct = q[2]  # 連続正解数

        self.progress_var.set(f"復習中... 残り問題: {len(available_questions)-1} (連続正解: {self.consecutive_correct})")
        self.word_label.config(text=self.current_question)
        self.speak_current_word()

        # 選択肢生成
        choices = self.master_app.db.get_random_word_choices(self.correct_answer, 3)
        all_choices = [self.correct_answer] + choices

        # 重複除去
        unique_choices = []
        seen = set()
        for choice in all_choices:
            normalized = choice.lower().replace('・', '').replace(' ', '').replace('、', '').replace('。', '')
            if normalized not in seen and choice.strip():
                unique_choices.append(choice)
                seen.add(normalized)

        # 4つ確保
        while len(unique_choices) < 4:
            dummy_options = ["該当なし", "不明", "その他", "関連語なし"]
            for dummy in dummy_options:
                if dummy not in unique_choices:
                    unique_choices.append(dummy)
                    break

        unique_choices = unique_choices[:4]
        random.shuffle(unique_choices)

        # ボタン作成
        for i, choice in enumerate(unique_choices):
            btn = ttk.Button(self.btn_frame, text=f"{chr(65+i)}. {choice}",
                             command=lambda c=choice: self.check_answer(c))
            btn.pack(fill=tk.X, pady=8, padx=50, ipady=10)

        self.current_q_index += 1

    def check_answer(self, choice):
        is_correct = choice == self.correct_answer

        if is_correct:
            self.score += 1
            self.feedback_var.set("✅ 正解！")
            sound_file = SOUND_FILES["correct"]
            # データベースの正解カウントを更新
            self.master_app.db.update_wrong_question_score('word_choice', self.current_question, True)
        else:
            self.feedback_var.set(f"❌ 不正解。正解は「{self.correct_answer}」")
            sound_file = SOUND_FILES["incorrect"]
            # データベースの正解カウントをリセット
            self.master_app.db.update_wrong_question_score('word_choice', self.current_question, False)

        # 効果音再生
        if os.path.exists(sound_file):
            threading.Thread(target=self._play_sound, args=(sound_file,), daemon=True).start()

        for btn in self.btn_frame.winfo_children():
            btn.config(state="disabled")
        self.next_btn.config(state="normal")

        if HAS_TTS and hasattr(self, 'speak_button'):
            self.speak_button.config(state="normal")

    def show_final_score(self):
        self.word_label.config(text="復習完了！")
        self.feedback_var.set(f"復習終了：{self.current_q_index - 1} 問取り組み、{self.score} 問正解")
        for btn in self.btn_frame.winfo_children():
            btn.destroy()

        self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")
        if HAS_TTS and hasattr(self, 'speak_button'):
            self.speak_button.config(state="disabled")

        # 統計更新
        self.master_app.update_stats()

    def on_closing(self):
        self.destroy()


# --- 間違い単語クイズ（リスニング形式）---
class WrongWordQuizListeningWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__()
        self.master_app = master
        self.questions = list(questions)  # [(question_content, correct_answer, consecutive_correct), ...]
        self.total = len(questions)
        self.current_idx = 0
        self.is_speaking = False
        self.score = 0
        self.used_questions = set()

        self.title("間違い復習：単語リスニングクイズ")
        self.geometry("740x460")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_label = ttk.Label(self, text="", font=("", 12))
        self.progress_label.pack(pady=10)

        self.word_label = ttk.Label(self, text="...", font=("", 24, "bold"), justify=tk.CENTER)
        self.word_label.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.meaning_label = ttk.Label(self, text="（答えは下に表示されます）", font=("", 14), justify=tk.CENTER)
        self.meaning_label.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        self.judge_feedback = ttk.Label(self, text="", font=("", 11), foreground="blue")
        self.judge_feedback.pack(pady=(0, 6))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn_frame, text="🔊 もう一度聞く", command=self.speak_current_word)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.answer_btn = ttk.Button(btn_frame, text="答えを見る", command=self.show_answer)
        self.answer_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.correct_btn = ttk.Button(btn_frame, text="〇 正解", command=lambda: self.judge(True), state="disabled")
        self.correct_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.wrong_btn = ttk.Button(btn_frame, text="× 不正解", command=lambda: self.judge(False), state="disabled")
        self.wrong_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn_frame, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_next_question(self):
        # 利用可能な問題をフィルタリング（まだ出題していない問題）
        available_questions = [q for i, q in enumerate(self.questions) if i not in self.used_questions]
        
        if not available_questions:
            self.word_label.config(text="復習完了！")
            self.meaning_label.config(text=f"復習終了：{self.current_idx - 1} 問取り組み、{self.score} 問正解")
            self.answer_btn.config(state="disabled")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.correct_btn.config(state="disabled")
            self.wrong_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")

            # 統計更新
            self.master_app.update_stats()
            return

        # ランダムに1つ選択
        q = random.choice(available_questions)
        # 使用済みに追加
        original_index = self.questions.index(q)
        self.used_questions.add(original_index)
        
        self.current_word = q[0]  # 英単語
        self.current_meaning = q[1]  # 正解
        self.consecutive_correct = q[2]  # 連続正解数

        self.progress_label.config(text=f"復習中... 残り問題: {len(available_questions)-1} (連続正解: {self.consecutive_correct})")
        self.word_label.config(text=self.current_word)
        self.meaning_label.config(text="（下に日本語の意味が表示されます）")
        self.judge_feedback.config(text="")

        self.answer_btn.config(state="normal")
        self.next_btn.config(state="disabled")
        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        if hasattr(self, 'speak_btn'):
            self.speak_btn.config(state="normal")

        if HAS_TTS:
            self.speak_current_word()

        self.current_idx += 1

    def show_answer(self):
        self.meaning_label.config(text=self.current_meaning)
        self.answer_btn.config(state="disabled")
        self.correct_btn.config(state="normal")
        self.wrong_btn.config(state="normal")
        self.judge_feedback.config(text="正解なら『〇 正解』、間違いなら『× 不正解』を押してください。")

    def judge(self, is_correct: bool):
        if is_correct:
            self.score += 1
            self.judge_feedback.config(text="✅ 正解として記録しました")
            sound_file = SOUND_FILES["correct"]
            # データベースの正解カウントを更新
            self.master_app.db.update_wrong_question_score('word_listening', self.current_word, True)
        else:
            self.judge_feedback.config(text="❌ 不正解として記録しました")
            sound_file = SOUND_FILES["incorrect"]
            # データベースの正解カウントをリセット
            self.master_app.db.update_wrong_question_score('word_listening', self.current_word, False)

        if os.path.exists(sound_file):
            self._play_sound(sound_file)

        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        self.next_btn.config(state="normal")

    def _play_sound(self, sound_path, duration_ms=1500):
        try:
            winmm = ctypes.windll.winmm
            alias = f"se_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(duration_ms, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"speech_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_word(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_word,), daemon=True).start()

    def on_closing(self):
        self.destroy()


# --- 間違い例文クイズ（リスニング形式）---
class WrongSentenceQuizWindow(tk.Toplevel):
    def __init__(self, master, questions):
        super().__init__()
        self.master_app = master
        self.questions = list(questions)  # [(question_content, correct_answer, consecutive_correct), ...]
        self.total = len(questions)
        self.current_idx = 0
        self.is_speaking = False
        self.score = 0
        self.used_questions = set()

        self.title("間違い復習：例文リスニングクイズ")
        self.geometry("760x480")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_next_question()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        self.progress_label = ttk.Label(self, text="", font=("", 12))
        self.progress_label.pack(pady=10)

        self.english_label = ttk.Label(self, text="...", font=("", 16, "bold"), wraplength=700, justify=tk.CENTER)
        self.english_label.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.japanese_label = ttk.Label(self, text="（答えは下に表示されます）", font=("", 12), wraplength=700, justify=tk.CENTER)
        self.japanese_label.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        self.judge_feedback = ttk.Label(self, text="", font=("", 11), foreground="blue")
        self.judge_feedback.pack(pady=(0, 6))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn_frame, text="🔊 もう一度聞く", command=self.speak_current_sentence)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.answer_btn = ttk.Button(btn_frame, text="答えを見る", command=self.show_answer)
        self.answer_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.correct_btn = ttk.Button(btn_frame, text="〇 正解", command=lambda: self.judge(True), state="disabled")
        self.correct_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.wrong_btn = ttk.Button(btn_frame, text="× 不正解", command=lambda: self.judge(False), state="disabled")
        self.wrong_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn_frame, text="次の問題へ", command=self.show_next_question, state="disabled")
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_next_question(self):
        # 利用可能な問題をフィルタリング（まだ出題していない問題）
        available_questions = [q for i, q in enumerate(self.questions) if i not in self.used_questions]
        
        if not available_questions:
            self.english_label.config(text="復習完了！")
            self.japanese_label.config(text=f"復習終了：{self.current_idx - 1} 問取り組み、{self.score} 問正解")
            self.answer_btn.config(state="disabled")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.correct_btn.config(state="disabled")
            self.wrong_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_closing, state="normal")

            # 統計更新
            self.master_app.update_stats()
            return

        # ランダムに1つ選択
        q = random.choice(available_questions)
        # 使用済みに追加
        original_index = self.questions.index(q)
        self.used_questions.add(original_index)
        
        self.current_english = q[0]  # 英文
        self.current_japanese = q[1]  # 正解
        self.consecutive_correct = q[2]  # 連続正解数

        self.progress_label.config(text=f"復習中... 残り問題: {len(available_questions)-1} (連続正解: {self.consecutive_correct})")
        self.english_label.config(text=self.current_english)
        self.japanese_label.config(text="（下に日本語訳が表示されます）")
        self.judge_feedback.config(text="")

        self.answer_btn.config(state="normal")
        self.next_btn.config(state="disabled")
        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        if hasattr(self, 'speak_btn'):
            self.speak_btn.config(state="normal")

        if HAS_TTS:
            self.speak_current_sentence()

        self.current_idx += 1

    def show_answer(self):
        self.japanese_label.config(text=self.current_japanese)
        self.answer_btn.config(state="disabled")
        self.correct_btn.config(state="normal")
        self.wrong_btn.config(state="normal")
        self.judge_feedback.config(text="正解なら『〇 正解』、間違いなら『× 不正解』を押してください。")

    def judge(self, is_correct: bool):
        if is_correct:
            self.score += 1
            self.judge_feedback.config(text="✅ 正解として記録しました")
            sound_file = SOUND_FILES["correct"]
            # データベースの正解カウントを更新
            self.master_app.db.update_wrong_question_score('sentence', self.current_english, True)
        else:
            self.judge_feedback.config(text="❌ 不正解として記録しました")
            sound_file = SOUND_FILES["incorrect"]
            # データベースの正解カウントをリセット
            self.master_app.db.update_wrong_question_score('sentence', self.current_english, False)

        if os.path.exists(sound_file):
            self._play_sound(sound_file)

        self.correct_btn.config(state="disabled")
        self.wrong_btn.config(state="disabled")
        self.next_btn.config(state="normal")

    def _play_sound(self, sound_path, duration_ms=1500):
        try:
            winmm = ctypes.windll.winmm
            alias = f"se_{random.randint(1000,9999)}"
            path_abs = os.path.abspath(sound_path)
            winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
            winmm.mciSendStringW(f'play {alias}', None, 0, None)
            self.after(duration_ms, lambda: winmm.mciSendStringW(f'close {alias}', None, 0, None))
        except Exception as e:
            print(f"Sound playback error: {e}")

    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"speech_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak_current_sentence(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_english,), daemon=True).start()

    def on_closing(self):
        self.destroy()


# --- 復習ウィンドウ（単語：間違いのみを周回） ---
class WordReviewWindow(tk.Toplevel):
    def __init__(self, master, wrong_items):
        super().__init__()
        self.master_quiz = master  # 呼び出し元（音声再生手持ちのためではなく参照用）
        self.items = list(wrong_items)  # [(word, meaning), ...]
        self.total = len(self.items)
        self.idx = 0
        self.is_speaking = False

        self.title("復習：間違えた単語")
        self.geometry("700x420")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_item()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        self.progress = ttk.Label(self, text="", font=("", 12))
        self.progress.pack(pady=10)

        self.word = ttk.Label(self, text="", font=("", 24, "bold"), justify=tk.CENTER)
        self.word.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.meaning = ttk.Label(self, text="（答えは下に表示されます）", font=("", 14), justify=tk.CENTER)
        self.meaning.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        btn = ttk.Frame(self)
        btn.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn, text="🔊 もう一度聞く", command=self.speak)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.reveal_btn = ttk.Button(btn, text="答えを見る", command=self.reveal)
        self.reveal_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn, text="次へ", command=self.next_item)
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_item(self):
        if not self.items:
            self.word.config(text="復習対象はありません。")
            self.meaning.config(text="")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.reveal_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_close)
            return

        if self.idx >= self.total:
            self.idx = 0  # ループさせたい場合
        w, m = self.items[self.idx]
        self.current_word, self.current_meaning = w, m
        self.progress.config(text=f"復習 {self.idx + 1} / {self.total}")
        self.word.config(text=w)
        self.meaning.config(text="（下に日本語の意味が表示されます）")

        if HAS_TTS:
            self.speak()

    def reveal(self):
        self.meaning.config(text=self.current_meaning)

    def next_item(self):
        self.idx += 1
        self.show_item()

    # --- TTS（読み上げだけ。効果音は無し） ---
    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"rvw_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_word,), daemon=True).start()

    def on_close(self):
        self.destroy()

# --- 復習ウィンドウ（例文：間違いのみを周回） ---
class SentenceReviewWindow(tk.Toplevel):
    def __init__(self, master, wrong_items):
        super().__init__()
        self.master_quiz = master
        self.items = list(wrong_items)  # [(english, japanese)]
        self.total = len(self.items)
        self.idx = 0
        self.is_speaking = False

        self.title("復習：間違えた例文")
        self.geometry("760x480")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.create_widgets()
        self.show_item()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        self.progress = ttk.Label(self, text="", font=("", 12))
        self.progress.pack(pady=10)

        self.eng = ttk.Label(self, text="", font=("", 16, "bold"), wraplength=700, justify=tk.CENTER)
        self.eng.pack(pady=10, padx=20, expand=True, fill=tk.BOTH)

        self.ja = ttk.Label(self, text="（答えは下に表示されます）", font=("", 12), wraplength=700, justify=tk.CENTER)
        self.ja.pack(pady=6, padx=20, expand=True, fill=tk.BOTH)

        btn = ttk.Frame(self)
        btn.pack(pady=10, fill=tk.X, padx=20)

        if HAS_TTS:
            self.speak_btn = ttk.Button(btn, text="🔊 もう一度聞く", command=self.speak)
            self.speak_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.reveal_btn = ttk.Button(btn, text="答えを見る", command=self.reveal)
        self.reveal_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

        self.next_btn = ttk.Button(btn, text="次へ", command=self.next_item)
        self.next_btn.pack(side=tk.LEFT, expand=True, padx=5, ipady=5)

    def show_item(self):
        if not self.items:
            self.eng.config(text="復習対象はありません。")
            self.ja.config(text="")
            if hasattr(self, 'speak_btn'):
                self.speak_btn.config(state="disabled")
            self.reveal_btn.config(state="disabled")
            self.next_btn.config(text="閉じる", command=self.on_close)
            return

        if self.idx >= self.total:
            self.idx = 0
        e, j = self.items[self.idx]
        self.current_english, self.current_japanese = e, j
        self.progress.config(text=f"復習 {self.idx + 1} / {self.total}")
        self.eng.config(text=e)
        self.ja.config(text="（下に日本語訳が表示されます）")

        if HAS_TTS:
            self.speak()

    def reveal(self):
        self.ja.config(text=self.current_japanese)

    def next_item(self):
        self.idx += 1
        self.show_item()

    def _speak_task(self, text):
        async def _task():
            try:
                self.is_speaking = True
                communicate = edge_tts.Communicate(text, TTS_VOICE)
                await communicate.save(AUDIO_FILE_PATH)
                winmm = ctypes.windll.winmm
                alias = f"rvw_{random.randint(1000,9999)}"
                path_abs = os.path.abspath(AUDIO_FILE_PATH)
                winmm.mciSendStringW(f'open "{path_abs}" type mpegvideo alias {alias}', None, 0, None)
                winmm.mciSendStringW(f'play {alias} wait', None, 0, None)
                winmm.mciSendStringW(f'close {alias}', None, 0, None)
            except Exception as e:
                print(f"TTS Error: {e}")
            finally:
                self.is_speaking = False
                if os.path.exists(AUDIO_FILE_PATH):
                    try:
                        os.remove(AUDIO_FILE_PATH)
                    except OSError:
                        pass
        try:
            loop = asyncio.get_running_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(_task())

    def speak(self):
        if HAS_TTS and not self.is_speaking:
            threading.Thread(target=self._speak_task, args=(self.current_english,), daemon=True).start()

    def on_close(self):
        self.destroy()

# --- ヘルパー関数 ---
def load_config():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(config_data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=2)

def main():
    try:
        config_data = load_config()
        app = App(config_data)
        app.mainloop()
    except Exception:
        error_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"\n--- Integrated App Error at {error_time} ---\n")
            traceback.print_exc(file=f)
        messagebox.showerror("重大なエラー", "予期せぬエラーが発生しました。\n詳細は error_log.txt を確認してください。")

if __name__ == "__main__":
    main()