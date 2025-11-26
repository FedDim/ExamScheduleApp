using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Group = ExamScheduleApp.Model.Group;

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
                MessageBox.Show("Нет файла подходящему к данному типу данных", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(AddTextBox.Text))
            {
                MessageBox.Show("Заполните поле!");
                return;
            }

            string data = AddTextBox.Text;
            SimpleDatabaseHelper dbHelper = new SimpleDatabaseHelper();

            try
            {
                switch (_dataType)
                {
                    case DataType.SUBJECT:
                        if (dbHelper.SubjectExists(data))
                        {
                            MessageBox.Show("Такая дисциплина уже существует!", "Информация",
                                          MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        Subject subject = new Subject
                        {
                            ShortName9 = data.Trim()
                            // Остальные поля заполнятся автоматически в методе AddSubject
                        };
                        dbHelper.AddSubject(subject);
                        break;

                    case DataType.GROUP:
                        if (!Regex.IsMatch(data.ToUpper().Trim(), @"^[А-ЯЁ]{2}-\d{2}$"))
                        {
                            MessageBox.Show("Введите корректное название группы! Пример: ДВ-45");
                            return;
                        }

                        data = data.ToUpper();
                        if (dbHelper.GroupExists(data))
                        {
                            MessageBox.Show("Такая группа уже существует!", "Информация",
                                          MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        Group group = new Group { Name = data.Trim() };
                        dbHelper.AddGroup(group);
                        break;

                    case DataType.TEACHER:
                        data = data.ToUpper();
                        if (dbHelper.TeacherExists(data))
                        {
                            MessageBox.Show("Такой преподаватель уже существует!", "Информация",
                                          MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        Teacher teacher = new Teacher
                        {
                            Name = data.Trim(),
                            Classroom = "",
                            AcademicBuilding = 0
                        };
                        dbHelper.AddTeacher(teacher);
                        break;
                }

                // Вызываем событие для обновления главной формы
                DataAdded?.Invoke(_dataType);

                // Очищаем поле для следующего ввода
                AddTextBox.Clear();

                MessageBox.Show("Данные успешно добавлены!", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
