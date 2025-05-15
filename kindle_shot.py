# kindle_shot.py

import pyautogui
import pygetwindow as gw
import time
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from ctypes import windll # DPIAwareness用

# --- 設定項目 (ここを調整してください) ---

# Kindleアプリのウィンドウタイトル（部分一致で動作することが多い）
# 以下のいずれかの行のコメントを外し、ご自身の環境で最も安定して特定できるものを選択してください。
# KINDLE_WINDOW_TITLE = "Kindle" 
KINDLE_WINDOW_TITLE = "Kindle for PC" # 一般的なタイトル
# KINDLE_WINDOW_TITLE = "Obsidian でつなげる情報管理術【完成版】" # 現在開いている本のタイトルが含まれる場合 (本が変わるとNG)
# 注意: gw.getWindowsWithTitle() は大文字・小文字を区別する場合があります。

# 本文エリアの相対的なオフセットとサイズ (Kindleウィンドウに対するピクセル単位)
# これはご自身のKindleアプリの表示に合わせて調整が必須です。
# スクリーンショットツールなどでピクセル数を計測してください。
CONTENT_MARGIN_LEFT = 220     # 例: 左側の余白 (要計測・調整)
CONTENT_MARGIN_TOP = 120      # 例: 上側の余白 (要計測・調整)
CONTENT_MARGIN_RIGHT = 40     # 例: 右側の余白 (要計測・調整)
CONTENT_MARGIN_BOTTOM = 40    # 例: 下側の余白 (要計測・調整)

# ページ送り後の待機時間 (秒) - ページ表示が完了するのに十分な時間を設定
WAIT_AFTER_PAGE_TURN = 1.5

# --- グローバル変数 ---
output_folder = ""
page_counter = 1
running = False # 実行中フラグ

# --- 関数定義 ---
def select_output_folder():
    """保存先フォルダを選択する関数"""
    global output_folder
    selected_path = filedialog.askdirectory()
    if selected_path:
        output_folder = selected_path
        folder_label.config(text=f"保存先: {output_folder}")
    else:
        folder_label.config(text="保存先: 未選択")

def get_kindle_window_and_region():
    """Kindleウィンドウを取得し、スクリーンショット領域を計算する"""
    try:
        # デバッグ: 利用可能なウィンドウタイトル一覧を表示
        # print("DEBUG: Available window titles:", gw.getAllTitles())

        kindle_windows = gw.getWindowsWithTitle(KINDLE_WINDOW_TITLE)
        if not kindle_windows:
            print(f"警告: タイトルに '{KINDLE_WINDOW_TITLE}' を含むウィンドウが見つかりません。")
            return None, None
        
        kindle_win = kindle_windows[0] # 最初に見つかったウィンドウを使用
        print(f"DEBUG: 発見したKindleウィンドウ: '{kindle_win.title}', 位置: ({kindle_win.left},{kindle_win.top}), サイズ: ({kindle_win.width}x{kindle_win.height})")


        win_x, win_y, win_width, win_height = kindle_win.left, kindle_win.top, kindle_win.width, kindle_win.height

        content_x = win_x + CONTENT_MARGIN_LEFT
        content_y = win_y + CONTENT_MARGIN_TOP
        content_width = win_width - CONTENT_MARGIN_LEFT - CONTENT_MARGIN_RIGHT
        content_height = win_height - CONTENT_MARGIN_TOP - CONTENT_MARGIN_BOTTOM

        if content_width <= 0 or content_height <= 0:
            print(f"警告: 計算されたコンテンツ領域の幅 ({content_width}) または高さ ({content_height}) が0以下です。マージン設定を確認してください。")
            return kindle_win, None # ウィンドウは返すが領域はNone
        
        calculated_region = (content_x, content_y, content_width, content_height)
        print(f"DEBUG: 計算されたスクリーンショット領域 (x,y,w,h): {calculated_region}")
        return kindle_win, calculated_region

    except Exception as e:
        print(f"ウィンドウ/領域取得中にエラー: {type(e).__name__} - {e}")
        return None, None

def start_screenshot():
    """スクリーンショット処理を開始する"""
    global page_counter, running, output_folder

    if not output_folder:
        messagebox.showwarning("注意", "保存先フォルダを選択してください。")
        return

    try:
        num_pages_str = num_pages_entry.get()
        if not num_pages_str.isdigit() or int(num_pages_str) <= 0:
            messagebox.showwarning("注意", "ページ数を正しく入力してください。")
            return
        num_pages_to_capture = int(num_pages_str)
    except ValueError:
        messagebox.showwarning("注意", "ページ数は数値で入力してください。")
        return

    # Kindleウィンドウと領域を最初に取得試行
    initial_kindle_win, initial_region = get_kindle_window_and_region()
    if not initial_kindle_win:
        messagebox.showerror("エラー", f"タイトルに '{KINDLE_WINDOW_TITLE}' を含むウィンドウが見つかりません。\nKindleアプリを起動しているか、ウィンドウタイトル設定を確認してください。")
        return
    if not initial_region:
        messagebox.showerror("エラー", "スクリーンショット領域の計算に失敗しました。\nマージン設定 (`CONTENT_MARGIN_...`) を確認してください。")
        return

    confirm_text = (f"Kindleウィンドウ: '{initial_kindle_win.title}'\n"
                    f"計算されたSショット領域 (x,y,w,h): {initial_region}\n"
                    f"{num_pages_to_capture}ページのSショットを開始しますか？\n\n"
                    "開始前にKindleアプリの目的のページを表示し、\n"
                    "カウントダウン中にKindleウィンドウをクリックしてアクティブにしてください。")
    confirm = messagebox.askyesno("確認", confirm_text)
    if not confirm:
        return

    running = True
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    status_label.config(text="実行準備中...")
    root.update_idletasks() # GUIを即時更新

    for i_countdown in range(3, 0, -1):
        status_label.config(text=f"開始まで {i_countdown} 秒... (Kindleウィンドウをアクティブに！)")
        root.update_idletasks()
        time.sleep(1)
    
    page_counter = 1 # 開始時にカウンターをリセット

    for i_loop in range(num_pages_to_capture):
        if not running: # 中断フラグチェック
            status_label.config(text="中断しました。")
            break
        
        status_label.config(text=f"撮影中: {page_counter}/{num_pages_to_capture} ページ目")
        root.update_idletasks()

        try:
            # ループの各反復でウィンドウと領域を再取得し、アクティブ化
            current_kindle_win, current_region = get_kindle_window_and_region()
            if not current_kindle_win:
                messagebox.showerror("エラー", "処理中にKindleウィンドウが見つからなくなりました。処理を中断します。")
                running = False; break
            if not current_region:
                 messagebox.showerror("エラー", "処理中にスクリーンショット領域の計算に失敗しました。処理を中断します。")
                 running = False; break

            if current_kindle_win.isMinimized:
                current_kindle_win.restore()
            current_kindle_win.activate() # ウィンドウをアクティブ化
            time.sleep(0.3) # アクティブ化と表示安定のための待機 (環境により調整)

            active_win_debug = pyautogui.getActiveWindow()
            print(f"DEBUG: スクリーンショット直前のアクティブウィンドウ: '{active_win_debug.title if active_win_debug else 'なし'}'")

            screenshot = pyautogui.screenshot(region=current_region)
            
            filename = f"kindle_page_{page_counter:03d}.png" #例: kindle_page_001.png
            filepath = os.path.join(output_folder, filename)
            screenshot.save(filepath)
            print(f"保存しました: {filepath}")

            page_counter += 1

            if i_loop < num_pages_to_capture - 1: # 最後のページではページ送りをしない
                # ページ送り前にもアクティブ化
                current_kindle_win.activate()
                time.sleep(0.1) 

                active_win_debug_before_press = pyautogui.getActiveWindow()
                print(f"DEBUG: ページ送り直前のアクティブウィンドウ: '{active_win_debug_before_press.title if active_win_debug_before_press else 'なし'}'")
                
                pyautogui.press('right') # ページ送り
                print("DEBUG: 右矢印キーを押しました (ページ送り)。")
                time.sleep(WAIT_AFTER_PAGE_TURN) # ページ表示完了待機
        
        except Exception as e:
            error_message = f"処理中に予期せぬエラーが発生しました: {type(e).__name__} - {e}"
            messagebox.showerror("エラー", error_message)
            print(f"詳細エラー: {error_message}") 
            running = False # エラー発生時は処理を停止
            break
    
    if running: # 正常にループ完了した場合
        status_label.config(text=f"完了: {num_pages_to_capture}ページ撮影しました。")
    
    running = False # 処理終了または中断
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)
    # page_counter は次回のためにリセット済み

def stop_screenshot():
    """スクリーンショット処理を中断する"""
    global running
    if running: 
        running = False # ループの中断フラグを立てる
        status_label.config(text="中断処理中...") # GUIにフィードバック
        print("中断シグナルを受け取りました。次のループ反復で停止します。")

# --- GUI設定 ---
root = tk.Tk()
root.title("Kindle スクリーンショットツール")

# フォルダ選択フレーム
folder_frame = tk.Frame(root)
folder_frame.pack(pady=10, padx=10, fill=tk.X)
tk.Button(folder_frame, text="保存先フォルダ選択", command=select_output_folder).pack(side=tk.LEFT)
folder_label = tk.Label(folder_frame, text="保存先: 未選択", anchor="w")
folder_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

# ページ数入力フレーム
pages_frame = tk.Frame(root)
pages_frame.pack(pady=5, padx=10, fill=tk.X)
tk.Label(pages_frame, text="撮影ページ数:").pack(side=tk.LEFT)
num_pages_entry = tk.Entry(pages_frame, width=10)
num_pages_entry.insert(0, "10") # デフォルト値
num_pages_entry.pack(side=tk.LEFT, padx=5)

# 操作ボタンフレーム
buttons_frame = tk.Frame(root)
buttons_frame.pack(pady=10)
start_button = tk.Button(buttons_frame, text="開始", command=start_screenshot, width=10)
start_button.pack(side=tk.LEFT, padx=5)
stop_button = tk.Button(buttons_frame, text="中断", command=stop_screenshot, width=10, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)

# ステータス表示ラベル
status_label = tk.Label(root, text="待機中", relief=tk.SUNKEN, anchor="w", bd=1) # bdで枠線
status_label.pack(fill=tk.X, padx=10, pady=5, ipady=5)

# 使い方説明フレーム
instructions_frame = tk.LabelFrame(root, text="使い方 と 注意点", padx=10, pady=10)
instructions_frame.pack(pady=10, padx=10, fill=tk.X)
instructions_text = (
    "1. Kindleアプリを起動し、スクリーンショットを開始したい最初のページを開いておきます。\n"
    "2. 「保存先フォルダ選択」で、画像を保存するフォルダを選びます。\n"
    "3. 「撮影ページ数」に必要なページ数を入力します。\n"
    "4. 「開始」ボタンを押すと、3秒のカウントダウン後に処理が始まります。\n"
    "   カウントダウン中にKindleアプリのウィンドウをクリックしてアクティブにしてください。\n"
    "5. 処理を途中で止めたい場合は「中断」ボタンを押します。\n\n"
    "注意:\n"
    "- スクリプト上部の `KINDLE_WINDOW_TITLE` 及び `CONTENT_MARGIN_LEFT` 等の値を、\n"
    "  ご自身のKindleアプリの表示に合わせて調整してください。この調整が最も重要です。\n"
    "  (ウィンドウ枠を除いた、純粋な本文が表示される領域を指定します)\n"
    "- Kindleアプリのウィンドウサイズはある程度固定して使用することを推奨します。\n"
    "- Windowsのディスプレイ拡大率が100%以外の場合、座標がズレる可能性があります。\n"
    "- 著作権に十分配慮し、私的利用の範囲を超えないようにしてください。"
)
tk.Label(instructions_frame, text=instructions_text, justify=tk.LEFT, anchor="w").pack(fill=tk.X)


# --- メイン処理 ---
if __name__ == "__main__":
    # WindowsでのDPIスケーリング対応 (GUIがぼやけるのを防ぐ)
    try:
        windll.shcore.SetProcessDpiAwareness(1)
        print("DPI Awareness set to 1 (System Aware)")
    except AttributeError:
        # 古いWindowsや他のOSでは windll.shcore がない場合がある
        try:
            windll.user32.SetProcessDPIAware()
            print("DPI Awareness set using user32.SetProcessDPIAware()")
        except AttributeError:
            print("DPI Awareness setting not available or failed.")
    except Exception as e:
        print(f"DPI Awareness setting error: {e}")

    root.mainloop()