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

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            Console.WriteLine("[Benkei] アプリケーション終了");
        }
    }
}
