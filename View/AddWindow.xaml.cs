using ExamScheduleApp.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace ExamScheduleApp
{
    /// <summary>
    /// Логика взаимодействия для AddWindow.xaml
    /// </summary>

    public enum DataType
    {
        NULL,
        SUBJECT,
        GROUP,
        TEACHER
    }

    public partial class AddWindow : Window
    {
        private DataType _dataType = DataType.NULL;

        public event Action<DataType> DataAdded;

        public AddWindow(DataType dataType)
        {
            InitializeComponent();

            _dataType = dataType;

            switch (_dataType)
            {
                case DataType.SUBJECT:
                    AddTextBlock.Text = $"Введите Предмет";
                    break;
                case DataType.GROUP:
                    AddTextBlock.Text = $"Введите Группу";
                    break;
                case DataType.TEACHER:
                    AddTextBlock.Text = $"Введите Преподавателя";
                    break;
                default:
                    AddTextBlock.Text = $"Введите NULL";
                    break;
            }

            Loaded += (s, e) => AddTextBox.Focus();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    AddButtonClick(sender, e);
                    break;
                case Key.Escape:
                    Close();
                    break;
            }
        }

        private void AddButtonClick(object sender, RoutedEventArgs e)
        {
            if (_dataType.Equals(DataType.NULL))
            {
                MessageBox.Show("Нет файла подходящему к данному типу данных", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(AddTextBox.Text))
            {
                MessageBox.Show("Заполните поле!");
                return;
            }

            string filePath = string.Empty;
            string data = AddTextBox.Text;

            switch (_dataType)
            {

                case DataType.SUBJECT:

                    filePath = FileHelper.GetFilePathFromData("disciplines.txt");

                    break;
                case DataType.GROUP:

                    if (!Regex.IsMatch(data.ToUpper().Trim(), @"^[А-ЯЁ]{2}-\d{2}$"))
                    {
                        MessageBox.Show("Введите корректное название группы! Пример: ДВ-45");
                        return;
                    }

                    data = data.ToUpper();
                    filePath = FileHelper.GetFilePathFromData("groups.txt");

                    break;
                case DataType.TEACHER:

                    data = data.ToUpper();
                    filePath = FileHelper.GetFilePathFromData("teachers.txt");

                    break;
            }

            try
            {
                List<string> lines = new List<string>();

                // Чтение существующих строк, если файл существует
                if (File.Exists(filePath))
                {
                    lines = File.ReadAllLines(filePath).ToList();
                }

                string trimmedData = data.Trim();

                if (!lines.Contains(trimmedData))
                {
                    lines.Add(trimmedData);
                    lines.Sort();

                    File.WriteAllLines(filePath, new HashSet<string>(lines));

                    // Вызываем событие для обновления главной формы
                    DataAdded?.Invoke(_dataType);

                    // Очищаем поле для следующего ввода
                    AddTextBox.Clear();

                    MessageBox.Show("Данные успешно добавлены!", "Успех",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("Такие данные уже существуют!", "Информация",
                                  MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с файлом: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }
    }
}
