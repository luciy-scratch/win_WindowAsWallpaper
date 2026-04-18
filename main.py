import os
import sys
import json
import time
import subprocess
import threading
import win32gui
import win32con
import win32api
import win32process
import ctypes
import pystray
from PIL import Image, ImageDraw

# デバッグ用のフラグを定義
switchConsoleVisible_inResidentMode = True # True=常駐モード中にコンソールウィンドウが非表示に設定される

class WindowAsWallpaper:
    def __init__(self, config_path):
        self.config_path = config_path
        self.child_processes = []
        self.worker_w = None
        self.icon = None
        self.running = True
        self.console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()

    def get_worker_w(self):
        """WorkerWの取得。Wallpaper Engineなどで既に生成されている場合はそれを検出し、なければ生成させる。"""
        def find_worker():
            target_hwnd = [None]
            def enum_windows_callback(hwnd, _):
                # SHELLDLL_DefView（デスクトップアイコン層）を持つWindowを探す
                shell_view = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
                if shell_view:
                    # その背後に隠れている兄弟ウィンドウ（WorkerW）を取得
                    # target_hwnd[0] = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
                    
                    # SHELLDLL_DefViewの兄弟/背後にあるWorkerWを探す(Wallpaper Engine対応?)
                    found = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
                    if not found:
                        found = win32gui.FindWindowEx(hwnd, 0, "WorkerW", None)
                    if found:
                        target_hwnd[0] = found
                return True
            win32gui.EnumWindows(enum_windows_callback, None)
            return target_hwnd[0]

        # 1. まず既存のWorkerWを探索
        self.worker_w = find_worker()
        
        if self.worker_w:
            print(f"既存の WorkerW を検出しました: {hex(self.worker_w)}")
        else:
            # 2. 見つからない場合のみ Progman にメッセージを送信して生成を促す
            print("WorkerW が見つからないため、新規生成をリクエストします...")
            progman = win32gui.FindWindow("Progman", None)
            win32gui.SendMessageTimeout(progman, 0x052C, 0, 0, win32con.SMTO_NORMAL, 1000)
            self.worker_w = find_worker()

        return self.worker_w

    def setup_window_style(self, hwnd):
        """タイトルバーや枠線を非表示にする"""
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        # キャプション、枠線、システムメニューなどを削除
        style &= ~win32con.WS_CAPTION
        style &= ~win32con.WS_THICKFRAME
        style &= ~win32con.WS_MINIMIZEBOX
        style &= ~win32con.WS_MAXIMIZEBOX
        style &= ~win32con.WS_SYSMENU
        style &= ~win32con.WS_VSCROLL
        style &= ~win32con.WS_HSCROLL
        
        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
        
        # win32guiにShowScrollBarは無いためctypesを使用してWin32 APIを直接呼ぶ
        ctypes.windll.user32.ShowScrollBar(hwnd, win32con.SB_BOTH, False)

    def position_window(self, hwnd, config):
        """
        指定されたモニタの4x2グリッドに合わせてウィンドウを配置する
        config: {monitor_index, x, y, w, h, include_taskbar}
        """
        monitors = win32api.EnumDisplayMonitors()
        if config['monitor'] >= len(monitors):
            return

        monitor_info = win32api.GetMonitorInfo(monitors[config['monitor']][0])
        # 1の場合はWorkArea(タスクバー除外)、0の場合はMonitorRect(全体)
        use_rect = monitor_info['Work'] if config.get('taskbar', 1) == 1 else monitor_info['Monitor']
        
        m_left, m_top, m_right, m_bottom = use_rect
        m_width = m_right - m_left
        m_height = m_bottom - m_top

        # 4x2グリッド計算
        unit_w = m_width // 4
        unit_h = m_height // 2

        target_x = m_left + (config['x'] * unit_w)
        target_y = m_top + (config['y'] * unit_h)
        target_w = config['w'] * unit_w
        target_h = config['h'] * unit_h

        win32gui.MoveWindow(hwnd, target_x, target_y, target_w, target_h, True)

    def find_window_for_process(self, pid, timeout_ms):
        """プロセスのウィンドウが生成されるまで待機して取得する"""
        start_time = time.time()
        while (time.time() - start_time) * 1000 < timeout_ms:
            def callback(hwnd, hwnds):
                if win32gui.IsWindowVisible(hwnd):
                    _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if found_pid == pid:
                        hwnds.append(hwnd)
                return True
            
            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            if hwnds:
                return hwnds[0]
            time.sleep(0.5)
        return None

    def run(self):
        print("Win_WindowAsWallpaper (WAW) 起動中...")
        
        # 1. WorkerWの準備
        if not self.get_worker_w():
            print("エラー: WorkerWの取得に失敗しました。")
            return

        # 2. 設定の読み込み
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception as e:
            print(f"エラー: 設定ファイルの読み込みに失敗しました。 {e}")
            return

        # 3. 各アプリケーションの起動と配置
        for item in settings:
            print(f"起動中: {item['path']}")
            try:
                proc = subprocess.Popen(item['path'] + " " + item.get('args', ''))
                self.child_processes.append(proc)
                
                # ウィンドウの出現を待機
                wait_ms = item.get('wait_ms', 2000)
                hwnd = self.find_window_for_process(proc.pid, wait_ms)
                
                if hwnd:
                    # WorkerWを親に設定
                    win32gui.SetParent(hwnd, self.worker_w)
                    # スタイル設定
                    self.setup_window_style(hwnd)
                    # 位置設定
                    self.position_window(hwnd, item)
                    print(f"配置完了: {hwnd}")
                    
                    # ウィンドウからフォーカスを外す（デスクトップにフォーカスを戻す）
                    shell_window = ctypes.windll.user32.GetShellWindow()
                    if shell_window:
                        try:
                            win32gui.SetForegroundWindow(shell_window)
                        except:
                            pass
                else:
                    print(f"警告: {item['path']} のウィンドウが見つかりませんでした。")

            except Exception as e:
                print(f"エラー: {item['path']} の起動に失敗しました。 {e}")

        print("全てのプロセスが配置されました。常駐モードに移行します。")
        self.stay_resident()

    def _create_element_icon(self):
        """トレイアイコン用の画像を生成する（青い背景に白い四角）"""
        width, height = 64, 64
        image = Image.new('RGB', (width, height), color=(31, 117, 204))
        draw = ImageDraw.Draw(image)
        # 中央に白い矩形を描画
        draw.rectangle([16, 16, 48, 48], fill=(255, 255, 255))
        return image

    def _on_exit_clicked(self, icon, item):
        """トレイメニューのExitがクリックされた際のコールバック"""
        self.cleanup()

    def stay_resident(self):
        """システムトレイアイコンを表示して待機する"""
        # コンソールを非表示にする
        if self.console_hwnd and switchConsoleVisible_inResidentMode:
            win32gui.ShowWindow(self.console_hwnd, win32con.SW_HIDE)

        menu = pystray.Menu(
            pystray.MenuItem("Exit", self._on_exit_clicked)
        )
        self.icon = pystray.Icon("WAW", self._create_element_icon(), "WindowAsWallpaper", menu)
        
        try:
            # icon.run() はブロッキング処理となり、メインループとして機能します
            self.icon.run()
        except KeyboardInterrupt:
            self.cleanup()

    def cleanup(self):
        """終了時に子プロセスを停止する"""
        # コンソールを再表示する
        if self.console_hwnd and switchConsoleVisible_inResidentMode:
            win32gui.ShowWindow(self.console_hwnd, win32con.SW_SHOW)

        print("終了処理中...")
        self.running = False
        if self.icon:
            self.icon.stop()
        for proc in self.child_processes:
            try:
                proc.terminate()
            except:
                pass
        print("終了しました。")

if __name__ == "__main__":
    # 簡易的な設定ファイルのデフォルトパス
    config_file = "settings.json"
    if len(sys.argv) > 1:
        config_file = sys.argv[1]

    if not os.path.exists(config_file):
        # サンプル設定ファイルの作成
        sample_config = [
            {
                "path": "conhost.exe",
                "args": "cmd.exe /k echo WindowAsWallpaper - conhost Sample",
                "monitor": 0,
                "x": 0, "y": 0, "w": 2, "h": 2,
                "wait_ms": 1000,
                "taskbar": 1
            }
        ]
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(sample_config, f, indent=4)
        print(f"サンプル設定ファイルを作成しました: {config_file}")

    waw = WindowAsWallpaper(config_file)
    
    # シャットダウン検知用のダミーウィンドウ作成（必要に応じて）
    # ここでは単純な実行として記述
    waw.run()