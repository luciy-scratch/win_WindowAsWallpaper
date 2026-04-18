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
import win32job
import ctypes
import pystray
from PIL import Image, ImageDraw
from rich.console import Console
from rich.rule import Rule

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
        self.console = Console(highlight=False)

        # Windows Job Object を使用して、起動したプロセスをグループ化し、
        # 親プロセス終了時に子プロセスも確実に終了するように設定する
        self.job_handle = win32job.CreateJobObject(None, "")
        extended_info = win32job.QueryInformationJobObject(
            self.job_handle, win32job.JobObjectExtendedLimitInformation
        )
        extended_info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(
            self.job_handle, win32job.JobObjectExtendedLimitInformation, extended_info
        )

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
            self.console.print(f"[bold green]既存の WorkerW を検出しました:[/bold green] [white]{hex(self.worker_w)}[/white]")
        else:
            # 2. 見つからない場合のみ Progman にメッセージを送信して生成を促す
            self.console.print("[yellow]WorkerW が見つからないため、新規生成をリクエストします...[/yellow]")
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

    def find_window_for_process(self, pid, exe_path, timeout_ms, target_title=None):
        """プロセスのウィンドウが生成されるまで待機して取得する"""
        start_time = time.time()
        target_exe_name = os.path.basename(exe_path).lower()

        while (time.time() - start_time) * 1000 < timeout_ms:
            def callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    # 1. ウィンドウタイトルでの一致確認 (設定がある場合、優先度高)
                    if target_title:
                        window_title = win32gui.GetWindowText(hwnd)
                        if target_title.lower() in window_title.lower():
                            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                            results.append((hwnd, found_pid))
                            return False

                    _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                    # 2. 直接のPID一致を確認
                    if found_pid == pid:
                        results.append((hwnd, found_pid))
                        return False
                    
                    # 3. PID不一致でも実行ファイル名が一致すれば採用（リダイレクト対策）
                    try:
                        h_proc = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, found_pid)
                        found_exe = win32process.GetModuleFileNameEx(h_proc, 0)
                        if os.path.basename(found_exe).lower() == target_exe_name:
                            results.append((hwnd, found_pid))
                            return False
                    except:
                        pass
                return True
            
            results = []
            win32gui.EnumWindows(callback, results)
            if results:
                return results[0]
            time.sleep(0.5)
        return None

    def run(self):
        self.console.print(Rule("[bold blue]Win_WindowAsWallpaper (WAW)[/bold blue]", style="blue", characters="="))
        self.console.print("[dim]Windows Desktop Enhancement Tool[/dim]", justify="center")
        self.console.print("")
        
        # 1. WorkerWの準備
        if not self.get_worker_w():
            self.console.print("[bold red]エラー:[/bold red] WorkerWの取得に失敗しました。")
            return

        # 2. 設定の読み込み
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
        except Exception as e:
            self.console.print(f"[bold red]エラー:[/bold red] 設定ファイルの読み込みに失敗しました。 [dim]{e}[/dim]")
            return

        # 3. 各アプリケーションの起動と配置
        for i, item in enumerate(settings):
            exe_name = os.path.basename(item['path'])
            self.console.print(f"🚀 [bold]({i+1}/{len(settings)}) 起動中:[/bold] [cyan]{exe_name}[/cyan] [dim]{item.get('args', '')}[/dim]")
            try:
                proc = subprocess.Popen(item['path'] + " " + item.get('args', ''))
                
                # ジョブオブジェクトにプロセスを割り当て
                try:
                    win32job.AssignProcessToJobObject(self.job_handle, proc._handle)
                except:
                    pass

                self.child_processes.append(proc)
                
                # ウィンドウの出現を待機
                wait_ms = item.get('wait_ms', 2000)
                result = self.find_window_for_process(proc.pid, item['path'], wait_ms, item.get('title'))
                
                if result:
                    hwnd, found_pid = result
                    # 真のウィンドウプロセスが別PID（リダイレクト等）の場合もジョブに追加
                    if found_pid != proc.pid:
                        try:
                            h_found = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, found_pid)
                            win32job.AssignProcessToJobObject(self.job_handle, h_found)
                        except:
                            pass

                    # WorkerWを親に設定
                    win32gui.SetParent(hwnd, self.worker_w)
                    # スタイル設定
                    self.setup_window_style(hwnd)
                    # 位置設定
                    self.position_window(hwnd, item)
                    self.console.print(f"   ∟ [bold green]配置完了:[/bold green] HWND:[white]{hwnd}[/white] | Monitor:{item['monitor']} | Grid:({item['x']},{item['y']}) Size:[white]{item['w']}x{item['h']}[/white]")
                    
                    # ウィンドウからフォーカスを外す（デスクトップにフォーカスを戻す）
                    shell_window = ctypes.windll.user32.GetShellWindow()
                    if shell_window:
                        try:
                            win32gui.SetForegroundWindow(shell_window)
                        except:
                            pass
                else:
                    self.console.print(f"   ∟ [bold yellow]警告:[/bold yellow] ウィンドウが見つかりませんでした。")

            except Exception as e:
                self.console.print(f"   ∟ [bold red]エラー:[/bold red] 起動に失敗しました。 [dim]{e}[/dim]")

        self.console.print("\n[bold green]全てのプロセスが配置されました。[/bold green]")
        self.console.print("[bold yellow]3秒後に常駐モードに移行します...[/bold yellow]")
        time.sleep(3)
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

        self.console.print("[bold red]終了処理中...[/bold red]")
        self.running = False
        if self.icon:
            self.icon.stop()
        # ジョブオブジェクトのハンドルを閉じることで、グループ化された全プロセスを終了させる
        if self.job_handle:
            self.job_handle.Close()
        self.console.print("[bold green]終了しました。3秒後に閉じます。[/bold green]")
        time.sleep(3)

if __name__ == "__main__":
    # 簡易的な設定ファイルのデフォルトパス
    config_file = "settings.json"
    if len(sys.argv) > 1:
        config_file = sys.argv[1]

    if not os.path.exists(config_file):
        # サンプル設定ファイルの作成
        sample_config = [
            {
                "path": os.path.join(os.environ.get('SystemRoot', 'C:\\Windows'), 'System32', 'conhost.exe'),
                "args": "cmd.exe /k \"title conhost Sample && echo WindowAsWallpaper - conhost Sample\"",
                "title": "conhost Sample",
                "monitor": 0,
                "x": 0, "y": 0, "w": 2, "h": 2,
                "wait_ms": 1000,
                "taskbar": 1
            }
        ]
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(sample_config, f, indent=4)
        Console(highlight=False).print(f"[bold green]サンプル設定ファイルを作成しました:[/bold green] {config_file}")

    waw = WindowAsWallpaper(config_file)
    
    # シャットダウン検知用のダミーウィンドウ作成（必要に応じて）
    # ここでは単純な実行として記述
    waw.run()