using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Benkei
{
    public partial class Form1 : Form
    {
        private KeyboardInterceptor _interceptor;
        private NotifyIcon _notifyIcon;
        private ContextMenuStrip _trayMenu;
        private bool _isExiting;

        public Form1()
        {
            InitializeComponent();
            InitializeTrayComponents();
            Load += OnLoaded;
            FormClosing += OnClosing;
            Resize += OnResized;
        }

        private void OnLoaded(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine("[Benkei] フォームロード開始");
                var configPath = ResolveConfigPath();
                Console.WriteLine($"[Benkei] 設定ファイル: {configPath}");
                var loader = new NaginataConfigLoader();
                var rules = loader.Load(configPath);
                Console.WriteLine($"[Benkei] ルール読み込み完了: {rules.Count}件");
                var engine = new NaginataEngine(rules);
                _interceptor = new KeyboardInterceptor(engine);
                _interceptor.Start();
                Console.WriteLine("[Benkei] キーボードフック開始");
                statusLabel.Text = $"動作中: {configPath}";
                HideToTray();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Benkei] エラー: {ex.Message}");
                Console.WriteLine($"[Benkei] スタックトレース: {ex.StackTrace}");
                statusLabel.Text = "エラー: 設定を読み込めませんでした";
                MessageBox.Show(this, ex.Message, "Benkei", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (!_isExiting && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                HideToTray();
                return;
            }

            Console.WriteLine("[Benkei] フォームクローズ開始");
            _interceptor?.Dispose();
            Console.WriteLine("[Benkei] キーボードフック停止");
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }
            _trayMenu?.Dispose();
        }

        private static string ResolveConfigPath()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidate = Path.Combine(baseDir, "Naginata.yaml");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            throw new FileNotFoundException("Naginata.yaml が見つかりません。", candidate);
        }

        private void InitializeTrayComponents()
        {
            _trayMenu = new ContextMenuStrip();
            var showMenuItem = new ToolStripMenuItem("状態を表示", null, OnTrayShowClick);
            var exitMenuItem = new ToolStripMenuItem("終了", null, OnTrayExitClick);
            _trayMenu.Items.Add(showMenuItem);
            _trayMenu.Items.Add(new ToolStripSeparator());
            _trayMenu.Items.Add(exitMenuItem);

            _notifyIcon = new NotifyIcon
            {
                Icon = Properties.Resources.NaginataTrayIcon,
                Text = "Benkei",
                Visible = false,
                ContextMenuStrip = _trayMenu
            };
            _notifyIcon.DoubleClick += OnTrayShowClick;
        }

        private void OnTrayShowClick(object sender, EventArgs e)
        {
            ShowFromTray();
        }

        private void OnTrayExitClick(object sender, EventArgs e)
        {
            ExitFromTray();
        }

        private void ExitFromTray()
        {
            _isExiting = true;
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
            }
            Close();
        }

        private void HideToTray()
        {
            if (_notifyIcon == null)
            {
                return;
            }

            _notifyIcon.Visible = true;
            if (WindowState != FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Minimized;
            }
            ShowInTaskbar = false;
            Hide();
        }

        private void ShowFromTray()
        {
            Show();
            ShowInTaskbar = true;
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void OnResized(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized && !_isExiting)
            {
                HideToTray();
            }
        }
    }
}
