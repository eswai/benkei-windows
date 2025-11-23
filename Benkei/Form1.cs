using System;
using System.IO;
using System.Windows.Forms;

namespace Benkei
{
    public partial class Form1 : Form
    {
        private KeyboardInterceptor _interceptor;

        public Form1()
        {
            InitializeComponent();
            Load += OnLoaded;
            FormClosing += OnClosing;
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
            Console.WriteLine("[Benkei] フォームクローズ開始");
            _interceptor?.Dispose();
            Console.WriteLine("[Benkei] キーボードフック停止");
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
    }
}
