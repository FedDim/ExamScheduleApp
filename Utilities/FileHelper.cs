using System;
using System.IO;
using System.Windows;

namespace ExamScheduleApp.Utilities
{
    public static class FileHelper
    {
        public static string GetFilePathFromData(string file)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", file);
            if (File.Exists(filePath)) return filePath;
            else
            {
                filePath = string.Empty;
                try
                {
                    string sourceFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "data", file));
                    string targerDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");

                    if (!Directory.Exists(targerDirectory)) Directory.CreateDirectory(targerDirectory);

                    filePath = Path.Combine(targerDirectory, file);

                    if (File.Exists(sourceFilePath))
                    {
                        File.Copy(sourceFilePath, filePath);
                    }
                    else
                    {
                        File.Create(filePath).Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка : {ex.Message}");
                }

                return filePath;
            }
        }
    }
}
