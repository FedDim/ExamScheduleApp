using System;
using System.IO;
using System.Windows;

namespace ExamScheduleApp.Utilities
{
    public static class Logger
    {
        private static readonly string LogFolder;
        private static readonly string CurrentLogPath;
        private static readonly object _lock = new object();

        static Logger()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                LogFolder = Path.Combine(baseDir, "Data", "Logs");
                Directory.CreateDirectory(LogFolder);

                CurrentLogPath = Path.Combine(LogFolder, "ExamSchedule.log");

                Log("=== Logger успешно инициализирован ===", "SYSTEM");
                Log($"Лог-файл создан по пути: {CurrentLogPath}", "SYSTEM");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка инициализации Logger:\n{ex.Message}",
                    "Критическая ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static void Log(string message, string level = "INFO")
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string logLine = $"[{timestamp}] [{level,-5}] {message}";

            lock (_lock)
            {
                try
                {
                    File.AppendAllText(CurrentLogPath, logLine + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось записать в лог:\n{ex.Message}\nПуть: {CurrentLogPath}");
                }
            }
        }

        public static void Info(string msg) => Log(msg, "INFO");
        public static void Warning(string msg) => Log(msg, "WARN");
        public static void Error(string msg, Exception ex = null)
        {
            string full = msg;
            if (ex != null) full += $" | {ex.Message}";
            Log(full, "ERROR");
        }
    }
}