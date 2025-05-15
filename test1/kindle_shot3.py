# kindle_shot3.py

import pyautogui
import pygetwindow as gw
import time
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from ctypes import windll

# --- 設定項目 (手動指定を使わない場合のデフォルト値) ---
KINDLE_WINDOW_TITLE = "Kindle for PC" 
CONTENT_MARGIN_LEFT = 220    
CONTENT_MARGIN_TOP = 120     
CONTENT_MARGIN_RIGHT = 40    
CONTENT_MARGIN_BOTTOM = 40   
WAIT_AFTER_PAGE_TURN = 1.5

# --- グローバル変数 ---
output_folder = ""
page_counter = 1
running = False
manual_region = None # 手動で指定されたスクリーンショット範囲 (x, y, w, h)

# --- 関数定義 ---
def select_output_folder():
    global output_folder
    selected_path = filedialog.askdirectory()
    if selected_path:
        output_folder = selected_path
        folder_label.config(text=f"保存先: {output_folder}")
    else:
        folder_label.config(text="保存先: 未選択")

def get_kindle_window_and_default_region():
    try:
        kindle_windows = gw.getWindowsWithTitle(KINDLE_WINDOW_TITLE)
        if not kindle_windows:
            print(f"警告: タイトルに '{KINDLE_WINDOW_TITLE}' を含むウィンドウが見つかりません。")
            return None, None
        
        kindle_win = kindle_windows[0]
        win_x, win_y, win_width, win_height = kindle_win.left, kindle_win.top, kindle_win.width, kindle_win.height

        content_x = win_x + CONTENT_MARGIN_LEFT
        content_y = win_y + CONTENT_MARGIN_TOP
        content_width = win_width - CONTENT_MARGIN_LEFT - CONTENT_MARGIN_RIGHT
        content_height = win_height - CONTENT_MARGIN_TOP - CONTENT_MARGIN_BOTTOM

        if content_width <= 0 or content_height <= 0:
            print(f"警告: デフォルト領域の幅/高さが0以下です。マージン設定を確認してください。")
            return kindle_win, None
        
        return kindle_win, (content_x, content_y, content_width, content_height)

    except Exception as e:
        print(f"デフォルト領域取得中にエラー: {type(e).__name__} - {e}")
        return None, None

def select_region_manually():
    global manual_region 

    root.withdraw() 
    time.sleep(0.3) 

    selector_win = tk.Toplevel(root)
    selector_win.attributes("-fullscreen", True)
    selector_win.attributes("-alpha", 0.3)    
    selector_win.attributes("-topmost", True) 
    selector_win.wait_visibility(selector_win)
    
    canvas = tk.Canvas(selector_win, cursor="cross", bg="grey")
    canvas.pack(fill=tk.BOTH, expand=True)

    rect_coords = {"x1": 0, "y1": 0, "x2": 0, "y2": 0}
    current_rect_item = None 

    def on_button_press(event):
        nonlocal current_rect_item 
        rect_coords["x1"] = event.x_root
        rect_coords["y1"] = event.y_root
        if current_rect_item:
            canvas.delete(current_rect_item)
        current_rect_item = canvas.create_rectangle(rect_coords["x1"], rect_coords["y1"], 
                                                    rect_coords["x1"], rect_coords["y1"], 
                                                    outline="red", width=2)
    def on_mouse_drag(event):
        nonlocal current_rect_item 
        rect_coords["x2"] = event.x_root
        rect_coords["y2"] = event.y_root
        if current_rect_item: 
            canvas.coords(current_rect_item, rect_coords["x1"], rect_coords["y1"], 
                                        rect_coords["x2"], rect_coords["y2"])
    def on_button_release(event):
        rect_coords["x2"] = event.x_root
        rect_coords["y2"] = event.y_root
        
        x = min(rect_coords["x1"], rect_coords["x2"])
        y = min(rect_coords["y1"], rect_coords["y2"])
        width = abs(rect_coords["x1"] - rect_coords["x2"])
        height = abs(rect_coords["y1"] - rect_coords["y2"])

        print(f"DEBUG (on_button_release): raw_coords: {rect_coords}")
        print(f"DEBUG (on_button_release): calculated_region: x={x}, y={y}, w={width}, h={height}")

        if width > 5 and height > 5: 
            manual_region = (x, y, width, height) 
            region_status_label.config(text=f"手動指定範囲: {manual_region}")
            print(f"DEBUG (on_button_release): 設定されたmanual_region: {manual_region}")
        else:
            manual_region = None 
            region_status_label.config(text="手動指定範囲: 未指定 (または小さすぎ)")
            print(f"DEBUG (on_button_release): manual_regionは設定されませんでした (width={width}, height={height})")

        selector_win.destroy()
        root.deiconify()

    canvas.bind("<ButtonPress-1>", on_button_press)
    canvas.bind("<B1-Motion>", on_mouse_drag)
    canvas.bind("<ButtonRelease-1>", on_button_release)

    messagebox.showinfo("範囲指定", "Kindleアプリなど、キャプチャしたいウィンドウの上で\nマウスをドラッグして範囲を指定してください。\n\n指定後、このメッセージボックスを閉じてください。", parent=selector_win)
    selector_win.focus_force()


def start_screenshot():
    global page_counter, running, output_folder, manual_region

    if not output_folder:
        messagebox.showwarning("注意", "保存先フォルダを選択してください。")
        return

    try:
        num_pages_str = num_pages_entry.get()
        num_pages_to_capture = int(num_pages_str)
        if num_pages_to_capture <= 0: raise ValueError
    except ValueError:
        messagebox.showwarning("注意", "ページ数を正の整数で入力してください。")
        return

    screenshot_region_to_use = None
    kindle_win_for_activation = None 
    is_manual_region_active = False
    region_source_msg = "" # 確認ダイアログ用のメッセージ

    print(f"DEBUG: start_screenshot開始時のmanual_region: {manual_region}") # ★追加デバッグ

    if manual_region and isinstance(manual_region, tuple) and len(manual_region) == 4:
        # manual_region が有効なタプル (x, y, w, h) であることを確認
        # さらに各要素が整数で、幅と高さが正であることも確認
        if (all(isinstance(n, int) for n in manual_region) and
                manual_region[2] > 0 and manual_region[3] > 0):
            screenshot_region_to_use = manual_region
            is_manual_region_active = True
            temp_win, _ = get_kindle_window_and_default_region()
            kindle_win_for_activation = temp_win 
            region_source_msg = f"手動指定された領域: {screenshot_region_to_use}"
            print(f"DEBUG: 手動範囲を使用します: {screenshot_region_to_use}")
        else:
            print(f"WARN: manual_regionの形式/値が無効です: {manual_region}。自動計算にフォールバックします。")
            manual_region = None # 無効な場合はクリアして自動計算へ
    
    if not is_manual_region_active: # 手動指定が無効または最初から無い場合
        kindle_win, default_region = get_kindle_window_and_default_region()
        if not kindle_win:
            messagebox.showerror("エラー", f"タイトルに '{KINDLE_WINDOW_TITLE}' を含むウィンドウが見つかりません。(手動指定も無効です)")
            return
        if not default_region:
            messagebox.showerror("エラー", "デフォルトのスクリーンショット領域計算に失敗。(手動指定も無効です)")
            return
        screenshot_region_to_use = default_region
        kindle_win_for_activation = kindle_win 
        region_source_msg = f"自動計算された領域 (Kindleウィンドウ基準): {screenshot_region_to_use}"
        print(f"DEBUG: 自動計算範囲を使用します: {screenshot_region_to_use}")

    if not screenshot_region_to_use: 
        messagebox.showerror("エラー", "スクリーンショット領域を決定できませんでした。")
        return

    confirm_text = (f"{region_source_msg}\n"
                    f"{num_pages_to_capture}ページのSショットを開始しますか？\n\n"
                    "開始前にKindleアプリの目的のページを表示し、\n"
                    "カウントダウン中にKindleウィンドウをクリックしてアクティブにしてください。")
    confirm = messagebox.askyesno("確認", confirm_text)
    if not confirm:
        return
    
    running = True
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    manual_region_button.config(state=tk.DISABLED)
    status_label.config(text="実行準備中...")
    root.update_idletasks()

    for i_countdown in range(3, 0, -1):
        status_label.config(text=f"開始まで {i_countdown} 秒... (Kindle/対象ウィンドウをアクティブに！)")
        root.update_idletasks()
        time.sleep(1)
    
    page_counter = 1

    for i_loop in range(num_pages_to_capture):
        if not running: 
            status_label.config(text="中断しました。")
            break
        
        status_label.config(text=f"撮影中: {page_counter}/{num_pages_to_capture} ページ目")
        root.update_idletasks()

        try:
            target_active_window_title_expected = "対象のアプリ" 
            if kindle_win_for_activation: 
                if kindle_win_for_activation.isMinimized:
                    kindle_win_for_activation.restore()
                kindle_win_for_activation.activate()
                target_active_window_title_expected = kindle_win_for_activation.title
                time.sleep(0.3) 
            elif is_manual_region_active:
                print("INFO: 手動範囲指定モードです。キャプチャ対象のウィンドウを前面にしてください。")
                time.sleep(0.3) 
            else:
                print("ERROR: アクティブ化すべきウィンドウが不明です。")
                time.sleep(0.3)

            active_win_debug = pyautogui.getActiveWindow()
            print(f"DEBUG: Sショット直前のアクティブウィンドウ: '{active_win_debug.title if active_win_debug else 'なし'}' (期待: '{target_active_window_title_expected}')")
            print(f"DEBUG: Sショットに使用する実際の領域: {screenshot_region_to_use}")

            if not (isinstance(screenshot_region_to_use, tuple) and len(screenshot_region_to_use) == 4 and
                    all(isinstance(n, int) for n in screenshot_region_to_use) and
                    screenshot_region_to_use[2] > 0 and screenshot_region_to_use[3] > 0):
                messagebox.showerror("致命的エラー", f"スクリーンショット領域の形式が無効です: {screenshot_region_to_use}\n処理を中断します。")
                running = False; break

            screenshot = pyautogui.screenshot(region=screenshot_region_to_use)
            
            filename = f"kindle_page_{page_counter:03d}.png"
            filepath = os.path.join(output_folder, filename)
            screenshot.save(filepath)
            print(f"保存しました: {filepath}")

            page_counter += 1

            if i_loop < num_pages_to_capture - 1:
                if kindle_win_for_activation: 
                    kindle_win_for_activation.activate()
                    time.sleep(0.1) 
                    pyautogui.press('right')
                    print("DEBUG: 右矢印キーを押しました (ページ送り)。")
                    time.sleep(WAIT_AFTER_PAGE_TURN)
                else:
                    print("WARN: Kindleウィンドウが特定できないため、自動ページ送りはスキップされました。手動でページを送ってください。")
                    time.sleep(WAIT_AFTER_PAGE_TURN)
        
        except Exception as e:
            error_message = f"処理中に予期せぬエラー: {type(e).__name__} - {e}"
            messagebox.showerror("エラー", error_message)
            print(f"詳細エラー: {error_message}") 
            running = False
            break
    
    if running: 
        status_label.config(text=f"完了: {num_pages_to_capture}ページ撮影しました。")
    
    running = False
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)
    manual_region_button.config(state=tk.NORMAL)

def stop_screenshot():
    global running
    if running: 
        running = False
        status_label.config(text="中断処理中...")
        print("中断シグナルを受け取りました。")

# --- GUI設定 (変更なし、ただしタイトルをV2.2に) ---
root = tk.Tk()
root.title("Kindle スクリーンショットツール V2.2") 

folder_frame = tk.Frame(root)
folder_frame.pack(pady=5, padx=10, fill=tk.X)
tk.Button(folder_frame, text="保存先フォルダ選択", command=select_output_folder).pack(side=tk.LEFT)
folder_label = tk.Label(folder_frame, text="保存先: 未選択", anchor="w")
folder_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

region_frame = tk.Frame(root)
region_frame.pack(pady=5, padx=10, fill=tk.X)
manual_region_button = tk.Button(region_frame, text="スクリーンショット範囲を手動指定", command=select_region_manually)
manual_region_button.pack(side=tk.LEFT)
region_status_label = tk.Label(region_frame, text="手動指定範囲: 未指定", anchor="w")
region_status_label.pack(side=tk.LEFT, padx=10)

pages_frame = tk.Frame(root)
pages_frame.pack(pady=5, padx=10, fill=tk.X)
tk.Label(pages_frame, text="撮影ページ数:").pack(side=tk.LEFT)
num_pages_entry = tk.Entry(pages_frame, width=10)
num_pages_entry.insert(0, "10") 
num_pages_entry.pack(side=tk.LEFT, padx=5)

buttons_frame = tk.Frame(root)
buttons_frame.pack(pady=10)
start_button = tk.Button(buttons_frame, text="開始", command=start_screenshot, width=10)
start_button.pack(side=tk.LEFT, padx=5)
stop_button = tk.Button(buttons_frame, text="中断", command=stop_screenshot, width=10, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)

status_label = tk.Label(root, text="待機中 (手動範囲指定がない場合は自動計算)", relief=tk.SUNKEN, anchor="w", bd=1)
status_label.pack(fill=tk.X, padx=10, pady=5, ipady=5)

instructions_frame = tk.LabelFrame(root, text="使い方 と 注意点", padx=10, pady=10)
instructions_frame.pack(pady=10, padx=10, fill=tk.X)
instructions_text = (
    "【範囲指定】:\n"
    "- 「スクリーンショット範囲を手動指定」ボタンで、マウスドラッグで範囲を指定できます。\n"
    "- 手動指定しない場合、スクリプト内の設定に基づいてKindleウィンドウから自動計算されます。\n"
    "【撮影手順】:\n"
    "1. Kindleアプリ等を起動し、キャプチャしたい最初のページを開きます。\n"
    "2. (任意) 「範囲を手動指定」で範囲を設定します。\n"
    "3. 「保存先フォルダ選択」で保存場所を選びます。\n"
    "4. 「撮影ページ数」を入力します。\n"
    "5. 「開始」ボタンを押し、カウントダウン中にキャプチャ対象ウィンドウをアクティブにします。\n\n"
    "【注意】:\n"
    "- 自動計算の場合、スクリプト上部の `KINDLE_WINDOW_TITLE` と `CONTENT_MARGIN_...` の調整が重要です。\n"
    "- 手動範囲指定の場合でも、自動ページ送りはKindleウィンドウが特定できた場合のみ機能します。"
)
tk.Label(instructions_frame, text=instructions_text, justify=tk.LEFT, anchor="w").pack(fill=tk.X)

if __name__ == "__main__":
    try:
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    
    root.mainloop()