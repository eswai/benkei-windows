using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Benkei
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Debug.WriteLine("[Benkei] アプリケーション起動");
            Debug.WriteLine($"[Benkei] 起動時刻: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

            Application.ThreadException += (sender, args) =>
            {
                LogException("UI thread", args.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                LogException("AppDomain", args.ExceptionObject as Exception);
            };
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                LogException("Main loop", ex);
                throw;
            }

            Debug.WriteLine("[Benkei] アプリケーション終了");
        }

        private static void LogException(string source, Exception ex)
        {
            try
            {
                if (ex == null)
                {
                    Debug.WriteLine($"[Benkei] {source}: (null exception)");
                    return;
                }

                var level = 0;
                var current = ex;
                while (current != null)
                {
                    var prefix = level == 0 ? source : $"{source}/Inner({level})";
                    Debug.WriteLine($"[Benkei] 例外({prefix}): {current.GetType().FullName}");
                    Debug.WriteLine($"[Benkei] メッセージ: {current.Message}");
                    Debug.WriteLine($"[Benkei] スタックトレース: {current.StackTrace}");
                    current = current.InnerException;
                    level++;
                }
            }
            catch (Exception logEx)
            {
                Debug.WriteLine($"[Benkei] 例外ログ出力に失敗 ({source}): {logEx.GetType().FullName} {logEx.Message}");
            }
        }
    }
}
