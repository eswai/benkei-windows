using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Benkei
{
    internal static class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);

        private const int AttachParentProcess = -1;

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // コンソールウィンドウを割り当て
            if (!AttachConsole(AttachParentProcess))
            {
                AllocConsole();
            }

            Console.WriteLine("[Benkei] アプリケーション起動");
            Console.WriteLine($"[Benkei] 起動時刻: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");

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

            Console.WriteLine("[Benkei] アプリケーション終了");
        }

        private static void LogException(string source, Exception ex)
        {
            try
            {
                if (ex == null)
                {
                    Console.WriteLine($"[Benkei] {source}: (null exception)");
                    return;
                }

                var level = 0;
                var current = ex;
                while (current != null)
                {
                    var prefix = level == 0 ? source : $"{source}/Inner({level})";
                    Console.WriteLine($"[Benkei] 例外({prefix}): {current.GetType().FullName}");
                    Console.WriteLine($"[Benkei] メッセージ: {current.Message}");
                    Console.WriteLine($"[Benkei] スタックトレース: {current.StackTrace}");
                    current = current.InnerException;
                    level++;
                }
            }
            catch (Exception logEx)
            {
                Console.WriteLine($"[Benkei] 例外ログ出力に失敗 ({source}): {logEx.GetType().FullName} {logEx.Message}");
            }
        }
    }
}
