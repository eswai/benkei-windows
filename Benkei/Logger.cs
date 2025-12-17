using System;

namespace Benkei
{
    internal static class Logger
    {
        private static Action<string> _sink = Console.WriteLine;

        public static void SetSink(Action<string> sink)
        {
            _sink = sink ?? Console.WriteLine;
        }

        public static void Log(string message)
        {
            _sink?.Invoke(message);
        }
    }
}
