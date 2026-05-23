using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.Linq;
using System.Windows;

namespace ExamScheduleApp.View
{
    public partial class DevWindow : Window
    {
        private ObservableCollection<ExamSchedule> _mainWindowExams;
        private ObservableCollection<ExamSchedule> _generatedExams;
        private SimpleDatabaseHelper _dbHelper;
        private Random _random;

        public DevWindow(ObservableCollection<ExamSchedule> mainWindowExams, SimpleDatabaseHelper dbHelper)
        {
            InitializeComponent();

            _mainWindowExams = mainWindowExams;
            _dbHelper = dbHelper;
            _generatedExams = new ObservableCollection<ExamSchedule>();
            _random = new Random();

            GeneratedExamsListBox.ItemsSource = _generatedExams;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(CountTextBox.Text, out int count) || count <= 0)
            {
                MessageBox.Show("Введите корректное число больше 0", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Получаем данные из базы
                var teachers = _dbHelper.GetTeachers();
                var subjects = _dbHelper.GetSubjects();
                var groups = _dbHelper.GetGroups();

                if (teachers.Count == 0 || subjects.Count == 0 || groups.Count == 0)
                {
                    MessageBox.Show("Недостаточно данных в базе для генерации", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Временные массивы времени и типа
                string[] times = { "9:00", "10:40", "11:00", "13:00", "13:30", "14:35", "15:00", "16:20" };
                string[] types = { "Экзамен", "Консультация" };
                string[] classrooms = { "101", "102", "103", "201", "202", "203", "301", "302", "303" };

                int generatedCount = 0;
                _generatedExams.Clear();

                for (int i = 0; i < count; i++)
                {
                    // Генерация случайных данных
                    var teacher1 = teachers[_random.Next(teachers.Count)];
                    var teacher2 = teachers.Count > 1 ? teachers[_random.Next(teachers.Count)] : null;
                    var subject = subjects[_random.Next(subjects.Count)];
                    var group = groups[_random.Next(groups.Count)];

                    // Генерация случайной даты (в течение месяца)
                    DateTime startDate = DateTime.Now;
                    DateTime randomDate = startDate.AddDays(_random.Next(30));
                    string dateString = randomDate.ToString("dd.MM.yyyy");

                    string time = times[_random.Next(times.Length)];
                    string type = types[_random.Next(types.Length)];
                    string classroom = classrooms[_random.Next(classrooms.Length)];

                    // Создание объекта экзамена
                    ExamSchedule exam = new ExamSchedule
                    {
                        Teacher1Id = teacher1.Id,
                        Teacher2Id = teacher2?.Id,
                        SubjectId = subject.Id,
                        GroupId = group.Id,
                        Classroom = classroom,
                        Teacher1Name = teacher1.Name,
                        Teacher2Name = teacher2?.Name,
                        SubjectName = subject.ShortName9,
                        GroupName = group.Name,
                        DepartmentName = group.Department,
                        ExamDate = dateString,
                        ExamTime = time,
                        ExamType = type
                    };

                    // Проверка на уникальность
                    if (!IsDuplicateExam(exam))
                    {
                        // Добавляем в базу данных (метод не меняем)
                        _dbHelper.AddExam(exam, false);

                        // Получаем ID добавленной записи через запрос последнего ID
                        int lastId = GetLastExamId();

                        if (lastId > 0)
                        {
                            exam.Id = lastId;

                            // Добавляем в локальную коллекцию
                            _generatedExams.Add(exam);
                            generatedCount++;
                        }
                    }
                    else
                    {
                        // Если дубликат, уменьшаем счетчик для повторной попытки
                        i--;
                    }
                }

                StatusTextBlock.Text = $"Сгенерировано {generatedCount} уникальных экзаменов";
                MessageBox.Show($"Успешно сгенерировано {generatedCount} экзаменов", "Генерация завершена",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при генерации экзаменов: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int GetLastExamId()
        {
            try
            {
                using (var connection = new SQLiteConnection(_dbHelper.GetLocalConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT seq FROM sqlite_sequence WHERE name='Exams'";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        var result = command.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) : 0;
                    }
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private bool IsDuplicateExam(ExamSchedule newExam)
        {
            // Проверяем в уже существующих экзаменах главного окна
            foreach (var existingExam in _mainWindowExams)
            {
                if (existingExam.Teacher1Id == newExam.Teacher1Id &&
                    existingExam.SubjectId == newExam.SubjectId &&
                    existingExam.GroupId == newExam.GroupId &&
                    existingExam.ExamDate == newExam.ExamDate &&
                    existingExam.ExamTime == newExam.ExamTime)
                {
                    return true;
                }
            }

            // Проверяем в сгенерированных в этой сессии
            foreach (var generatedExam in _generatedExams)
            {
                if (generatedExam.Teacher1Id == newExam.Teacher1Id &&
                    generatedExam.SubjectId == newExam.SubjectId &&
                    generatedExam.GroupId == newExam.GroupId &&
                    generatedExam.ExamDate == newExam.ExamDate &&
                    generatedExam.ExamTime == newExam.ExamTime)
                {
                    return true;
                }
            }

            return false;
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (_generatedExams.Count == 0) return;

            if (MessageBox.Show("Удалить сгенерированные экзамены из базы данных?", "Подтверждение",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    foreach (var exam in _generatedExams)
                    {
                        _dbHelper.DeleteExam(exam.Id, false);
                    }

                    _generatedExams.Clear();
                    StatusTextBlock.Text = "Сгенерированные экзамены удалены";
                    MessageBox.Show("Сгенерированные экзамены удалены из базы данных", "Информация",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // При закрытии окна проверяем, есть ли сгенерированные экзамены
            if (_generatedExams.Count > 0)
            {
                var result = MessageBox.Show("Сохранить сгенерированные экзамены и передать их в главное окно?",
                    "Сохранение данных", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // Добавляем все сгенерированные экзамены в основную коллекцию
                        foreach (var exam in _generatedExams)
                        {
                            // Проверяем, нет ли уже такого экзамена в основной коллекции
                            if (!_mainWindowExams.Any(examLamda => examLamda.Id == examLamda.Id))
                            {
                                _mainWindowExams.Add(exam);
                            }
                        }
                        MessageBox.Show($"Добавлено {_generatedExams.Count} экзаменов в расписание",
                            "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                        break;

                    case MessageBoxResult.No:
                        // Удаляем все сгенерированные экзамены из базы данных
                        foreach (var exam in _generatedExams)
                        {
                            try
                            {
                                _dbHelper.DeleteExam(exam.Id, false);
                            }
                            catch { }
                        }
                        MessageBox.Show("Сгенерированные экзамены удалены", "Информация",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                        break;

                    case MessageBoxResult.Cancel:
                        // Отменяем закрытие окна
                        e.Cancel = true;
                        return;
                }
            }
        }
    }
}