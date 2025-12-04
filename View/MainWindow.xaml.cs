using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using ExamScheduleApp.View;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExamScheduleApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<Teacher> _teachers;
        private List<Subject> _subjects;
        private List<Group> _groups;
        private ObservableCollection<ExamSchedule> _exams;
        public SimpleDatabaseHelper dbhelper = new SimpleDatabaseHelper();
        private bool _ctrlPressed = false, _tildePressed = false, _tabPressed = false, _isDevmode = true;

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                // Диагностика базы данных
                dbhelper.CheckDatabaseStructure();

                // Обновляем существующие записи
                dbhelper.UpdateExamFields();

                // Очищаем проблемные данные (для SimpleDatabaseHelper.cs)
                dbhelper.CleanProblematicData();

                // Очищаем все экзамены при запуске
                //dbhelper.ClearAllExams();

                // Загрузка данных
                LoadDataFromDatabase();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при запуске приложения: {ex.Message}");
            }
        }

        private void LoadDataFromDatabase()
        {
            try
            {
                // Загрузка данных из базы данных
                _teachers = dbhelper.GetTeachers();
                _subjects = dbhelper.GetSubjects();
                _groups = dbhelper.GetGroups();

                // Проверяем, что данные загружены
                if (_teachers == null || _teachers.Count == 0)
                {
                    MessageBox.Show("Не удалось загрузить преподавателей. Возможно, таблица пуста.");
                    _teachers = new List<Teacher>();
                }

                if (_subjects == null || _subjects.Count == 0)
                {
                    MessageBox.Show("Не удалось загрузить дисциплины. Возможно, таблица пуста.");
                    _subjects = new List<Subject>();
                }

                if (_groups == null || _groups.Count == 0)
                {
                    MessageBox.Show("Не удалось загрузить группы. Возможно, таблица пуста.");
                    _groups = new List<Group>();
                }

                // Заполнение ComboBox'ов
                cbTeachers.ItemsSource = _teachers;
                cbTeachers.DisplayMemberPath = "Name";
                cbTeachers.SelectedValuePath = "Id";

                SecondTeacherCB.ItemsSource = _teachers;
                SecondTeacherCB.DisplayMemberPath = "Name";
                SecondTeacherCB.SelectedValuePath = "Id";

                cbSubjects.ItemsSource = _subjects;
                cbSubjects.DisplayMemberPath = "ShortName9";
                cbSubjects.SelectedValuePath = "Id";

                GroupComboBox.ItemsSource = _groups;
                GroupComboBox.DisplayMemberPath = "Name";
                GroupComboBox.SelectedValuePath = "Id";

                // Обновление ComboBox'ов
                cbTeachers.ItemsSource = _teachers;
                SecondTeacherCB.ItemsSource = _teachers;
                cbSubjects.ItemsSource = _subjects;
                GroupComboBox.ItemsSource = _groups;

                // Инициализация коллекции экзаменов
                if (_exams == null)
                {
                    _exams = new ObservableCollection<ExamSchedule>();
                }

                // Загружаем существующие экзамены из базы
                RefreshExamsData();

                // Если в базе нет экзаменов, спрашиваем о загрузке из буфера
                if (_exams.Count == 0)
                {
                    var result = MessageBox.Show("В базе данных нет сохраненных экзаменов. Хотите загрузить экзамены из файла буфера?",
                        "Загрузка данных", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        LoadBuffer_Click(null, null);
                    }
                }
                else if (_isDevmode)
                {
                    // Показываем информационное сообщение
                    MessageBox.Show($"Загружено {_exams.Count} экзаменов из базы данных",
                        "Загрузка данных", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных из базы: {ex.Message}");
            }
        }
        private void AddExam(object sender, RoutedEventArgs e)
        {

            // Получаем объекты
            var teacher1 = (Teacher)cbTeachers.SelectedItem;
            var teacher2 = (Teacher)SecondTeacherCB.SelectedItem;
            var subject = (Subject)cbSubjects.SelectedItem;
            var group = (Group)GroupComboBox.SelectedItem;

            string surname = teacher1?.Name;
            string secondSurname = teacher2?.Name;

            // Дата, время и тип вводятся вручную
            DateTime? date = dpExamDate.SelectedDate;
            string dateString = date?.ToString("dd.MM.yyyy") ?? string.Empty;
            string time = TimeComboBox.SelectedItem != null ? TimeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;
            string type = TypeComboBox.SelectedItem != null ? TypeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;

            string subjectName = subject?.Name;
            string groupName = group?.Name;
            string cabinet = txtClassroom?.Text;

            string checkResult = CheckEnteredFields(surname, secondSurname, dateString, subjectName, groupName, time, cabinet, type);

            if (!checkResult.Equals(string.Empty))
            {
                MessageBox.Show(checkResult);
                return;
            }

            try
            {
                // Создание объекта для базы данных
                ExamSchedule exam = new ExamSchedule
                {
                    Teacher1Id = teacher1.Id,
                    Teacher2Id = teacher2?.Id,
                    SubjectId = subject.Id,
                    GroupId = group.Id,
                    Classroom = cabinet,
                    Teacher1Name = surname,
                    Teacher2Name = secondSurname,
                    SubjectName = subjectName,
                    GroupName = groupName,
                    DepartmentName = group.Department,
                    ExamDate = dateString,
                    ExamTime = time,
                    ExamType = type
                };

                // Добавление в базу данных
                dbhelper.AddExam(exam);

                // Получаем ID добавленной записи
                var addedExams = dbhelper.GetExamSchedule();
                if (addedExams.Count > 0)
                {
                    exam.Id = addedExams[addedExams.Count - 1].Id;
                }

                // Обновление интерфейса
                _exams.Add(exam);
                //ExamsDataGrid.Items.Refresh();

                // Очистка полей
                txtClassroom.Clear();
                dpExamDate.SelectedDate = null;

                MessageBox.Show("Экзамен успешно добавлен!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении экзамена: {ex.Message}");
            }
        }

        private string CheckEnteredFields(string surname, string secondSurname, string date,
                                string subject, string group, string time, string cabinet, string type)
        {
            string checkResult = string.Empty;

            if (string.IsNullOrEmpty(surname)) checkResult += "Не выбрана Фамилия первого преподавателя\n";
            //if (string.IsNullOrEmpty(secondSurname)) checkResult += "Не выбрана Фамилия второго преподавателя\n";
            if (string.IsNullOrEmpty(date)) checkResult += "Не выбрана дата проведения\n";
            if (string.IsNullOrEmpty(subject)) checkResult += "Не выбрана Дисциплина\n";
            if (string.IsNullOrEmpty(time)) checkResult += "Не заполнено Время\n";
            if (string.IsNullOrEmpty(cabinet)) checkResult += "Не заполнен Номер Кабинета\n";
            if (string.IsNullOrEmpty(type)) checkResult += "Не выбран Тип\n";

            if (!checkResult.Equals(string.Empty))
                checkResult = "Не все данные заполнены: \n" + checkResult;

            return checkResult;
        }

        #region Добавление Данных в Файлы Data
        private void AddTeacher_Click(object sender, RoutedEventArgs e) => ShowAddWindow(DataType.TEACHER);
        private void AddSubject_Click(object sender, RoutedEventArgs e) => ShowAddWindow(DataType.SUBJECT);
        private void AddGroup_Click(object sender, RoutedEventArgs e) => ShowAddWindow(DataType.GROUP);
        private void ShowAddWindow(DataType dataType)
        {
            AddWindow addWindow = new AddWindow(dataType);
            addWindow.DataAdded += (type) => RefreshData(type);
            addWindow.ShowDialog();
        }
        private void RefreshData(DataType dataType)
        {
            switch (dataType)
            {
                case DataType.TEACHER:
                    _teachers = dbhelper.GetTeachers();
                    cbTeachers.ItemsSource = _teachers;
                    SecondTeacherCB.ItemsSource = _teachers;
                    break;
                case DataType.SUBJECT:
                    _subjects = dbhelper.GetSubjects();
                    cbSubjects.ItemsSource = _subjects;
                    break;
                case DataType.GROUP:
                    _groups = dbhelper.GetGroups();
                    GroupComboBox.ItemsSource = _groups;
                    break;
            }
        }
        #endregion

        private void EditTeachers_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.TEACHER);
            window.Closed += (s, args) => RefreshData(DataType.TEACHER); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditSubjects_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.SUBJECT);
            window.Closed += (s, args) => RefreshData(DataType.SUBJECT); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditGroups_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.GROUP);
            window.Closed += (s, args) => RefreshData(DataType.GROUP); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void ShowScheduleTable_Click(object sender, RoutedEventArgs e)
        {
            var scheduleTableWindow = new ScheduleTableWindow(_exams);
            scheduleTableWindow.Owner = this;
            scheduleTableWindow.Closed += ScheduleTableWindow_Closed;
            scheduleTableWindow.ShowDialog();
        }

        private void ScheduleTableWindow_Closed(object sender, EventArgs e)
        {
            var scheduleTableWindow = sender as ScheduleTableWindow;

            if (scheduleTableWindow?.ExamsToDelete?.Count > 0)
            {
                try
                {
                    // Удаляем экзамены из базы данных
                    DeleteExamsFromDatabase(scheduleTableWindow.ExamsToDelete);

                    // Обновляем данные
                    RefreshExamsData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении экзаменов: {ex.Message}");
                }
            }
        }

        private void DeleteExamsFromDatabase(List<int> examIds)
        {
            foreach (var examId in examIds)
            {
                dbhelper.DeleteExam(examId);
            }
        }

        private void RefreshExamsData()
        {
            try
            {
                // Обновляем коллекцию экзаменов из базы данных
                var examList = dbhelper.GetExamSchedule();
                _exams.Clear();
                foreach (var exam in examList)
                {
                    _exams.Add(exam);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных: {ex.Message}");
            }
        }

        private void Window_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.LeftCtrl:
                    _ctrlPressed = true;
                    break;
                case Key.OemTilde:
                    _tildePressed = true;
                    break;
                case Key.Space:
                    _tabPressed = true;
                    break;
            }

            if (_isDevmode) CheckCombination();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.LeftCtrl:
                    _ctrlPressed = false;
                    break;
                case Key.OemTilde:
                    _tildePressed = false;
                    break;
                case Key.Tab:
                    _tabPressed = false;
                    break;
            }
        }

        private void CheckCombination()
        {
            if (_ctrlPressed && _tildePressed && _tabPressed && _isDevmode)
            {
                var window = new DevWindow(_exams, dbhelper)
                {
                    Owner = this
                };
                window.Closed += DevWindow_Closed;
                window.ShowDialog();

                _ctrlPressed = false;
                _tildePressed = false;
                _tabPressed = false;
            }
        }

        private void DevWindow_Closed(object sender, EventArgs e)
        {
            // Обновляем данные из базы после закрытия окна разработчика
            RefreshExamsData();
            MessageBox.Show("Расписание обновлено", "Информация",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #region Буфер
        private void SaveBuffer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";
                saveFileDialog.FileName = $"Данные экзаменов {DateTime.Now:dd.MM.yyyy HH-mm}.txt";
                saveFileDialog.Title = "Сохранить буфер экзаменов";

                if (saveFileDialog.ShowDialog() == true)
                {
                    SaveBufferToFile(saveFileDialog.FileName);
                    MessageBox.Show($"Буфер сохранен в файл:\n{saveFileDialog.FileName}",
                        "Сохранение завершено", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении буфера: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveBufferToFile(string filePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    // Заголовок с информацией
                    writer.WriteLine($"# Буфер экзаменов от {DateTime.Now:dd.MM.yyyy HH:mm}");
                    writer.WriteLine($"# Всего записей: {_exams.Count}");
                    writer.WriteLine("# Формат: Преподаватель1|Преподаватель2|Дата|Дисциплина|Группа|Кафедра|Время|Аудитория|Тип");
                    writer.WriteLine();

                    // Данные экзаменов
                    foreach (var exam in _exams)
                    {
                        string line = $"{exam.Teacher1Name ?? ""}|" +
                                     $"{exam.Teacher2Name ?? ""}|" +
                                     $"{exam.ExamDate}|" +
                                     $"{exam.SubjectName}|" +
                                     $"{exam.GroupName}|" +
                                     $"{exam.DepartmentName ?? ""}|" +
                                     $"{exam.ExamTime}|" +
                                     $"{exam.Classroom}|" +
                                     $"{exam.ExamType}";
                        writer.WriteLine(line);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка записи в файл: {ex.Message}");
            }
        }

        private void LoadBuffer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";
                openFileDialog.Title = "Выберите файл буфера для загрузки";

                if (openFileDialog.ShowDialog() == true)
                {
                    LoadFromBuffer(openFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при выборе файла: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadFromBuffer(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show($"Файл буфера не найден: {filePath}");
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
                int loadedCount = 0;
                int errorCount = 0;
                int skippedCount = 0;

                // Список для сбора ошибок
                List<string> errors = new List<string>();

                // Пропускаем заголовки (строки, начинающиеся с #)
                var dataLines = lines.Where(line => !string.IsNullOrWhiteSpace(line) && !line.TrimStart().StartsWith("#")).ToArray();

                if (dataLines.Length == 0)
                {
                    MessageBox.Show("Файл буфера не содержит данных для загрузки", "Информация",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Запрашиваем подтверждение загрузки
                var confirmResult = MessageBox.Show($"В файле найдено {dataLines.Length} строк данных. Начать загрузку?",
                    "Подтверждение загрузки", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirmResult != MessageBoxResult.Yes)
                    return;

                // Обрабатываем каждую строку
                for (int i = 0; i < dataLines.Length; i++)
                {
                    try
                    {
                        string line = dataLines[i];
                        string[] parts = line.Split('|');

                        // Проверяем количество частей
                        if (parts.Length != 9)
                        {
                            errorCount++;
                            errors.Add($"Строка {i + 1}: Неверный формат (ожидается 9 полей, получено {parts.Length})");
                            continue;
                        }

                        // Очищаем данные
                        for (int j = 0; j < parts.Length; j++)
                        {
                            parts[j] = parts[j].Trim();
                        }

                        // Парсим поля
                        string teacher1Name = parts[0];
                        string teacher2Name = parts[1];
                        string examDate = parts[2];
                        string subjectName = parts[3];
                        string groupName = parts[4];
                        string time = parts[6];
                        string classroom = parts[7];
                        string type = parts[8];

                        // Проверяем обязательные поля
                        if (string.IsNullOrEmpty(teacher1Name) ||
                            string.IsNullOrEmpty(subjectName) ||
                            string.IsNullOrEmpty(groupName) ||
                            string.IsNullOrEmpty(examDate) ||
                            string.IsNullOrEmpty(time) ||
                            string.IsNullOrEmpty(classroom) ||
                            string.IsNullOrEmpty(type))
                        {
                            errorCount++;
                            errors.Add($"Строка {i + 1}: Отсутствуют обязательные поля");
                            continue;
                        }

                        // Ищем преподавателей, дисциплину и группу в БД
                        var teacher1 = _teachers.FirstOrDefault(t => t.Name == teacher1Name);
                        var teacher2 = string.IsNullOrEmpty(teacher2Name) ? null : _teachers.FirstOrDefault(t => t.Name == teacher2Name);
                        var subject = _subjects.FirstOrDefault(s => s.ShortName9 == subjectName);
                        var group = _groups.FirstOrDefault(g => g.Name == groupName);

                        // Проверяем наличие необходимых данных
                        if (teacher1 == null)
                        {
                            errorCount++;
                            errors.Add($"Строка {i + 1}: Преподаватель '{teacher1Name}' не найден в базе данных");
                            continue;
                        }

                        if (subject == null)
                        {
                            errorCount++;
                            errors.Add($"Строка {i + 1}: Дисциплина '{subjectName}' не найдена в базе данных");
                            continue;
                        }

                        if (group == null)
                        {
                            errorCount++;
                            errors.Add($"Строка {i + 1}: Группа '{groupName}' не найдена в базе данных");
                            continue;
                        }

                        // Проверяем, нет ли такого экзамена уже в БД
                        bool exists = _exams.Any(e =>
                            e.Teacher1Id == teacher1.Id &&
                            e.SubjectId == subject.Id &&
                            e.GroupId == group.Id &&
                            e.ExamDate == examDate &&
                            e.ExamTime == time &&
                            e.Classroom == classroom &&
                            e.ExamType == type);

                        if (exists)
                        {
                            skippedCount++;
                            continue;
                        }

                        // Создаем новый экзамен
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
                            ExamDate = examDate,
                            ExamTime = time,
                            ExamType = type
                        };

                        // Добавляем в базу данных
                        dbhelper.AddExam(exam, false);
                        loadedCount++;

                        // Добавляем в локальную коллекцию
                        _exams.Add(exam);
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        errors.Add($"Строка {i + 1}: Ошибка обработки - {ex.Message}");
                    }
                }

                // Показываем итоговый отчет
                StringBuilder report = new StringBuilder();
                report.AppendLine("Загрузка завершена");
                report.AppendLine($"Всего строк в файле: {dataLines.Length}");
                report.AppendLine($"Успешно загружено: {loadedCount}");
                report.AppendLine($"Пропущено (дубликаты): {skippedCount}");
                report.AppendLine($"Ошибок: {errorCount}");

                // Если есть ошибки, показываем первые 10
                if (errors.Count > 0)
                {
                    report.AppendLine("\nПервые 10 ошибок:");
                    int showCount = Math.Min(errors.Count, 10);
                    for (int i = 0; i < showCount; i++)
                    {
                        report.AppendLine($"  {errors[i]}");
                    }

                    if (errors.Count > 10)
                    {
                        report.AppendLine($"  ... и еще {errors.Count - 10} ошибок");
                    }
                }

                MessageBox.Show(report.ToString(), "Результат загрузки",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке буфера: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearBuffer_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Проверяем, есть ли что очищать
                if (_exams.Count == 0)
                {
                    MessageBox.Show("Буфер пуст, нечего очищать.", "Информация",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Предлагаем сохранить буфер перед очисткой
                var saveResult = MessageBox.Show($"В буфере находится {_exams.Count} экзаменов.\n\nСохранить буфер в файл перед очисткой?",
                    "Сохранение буфера", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (saveResult == MessageBoxResult.Cancel)
                    return;

                if (saveResult == MessageBoxResult.Yes)
                {
                    SaveBuffer_Click(sender, e);
                }

                // Подтверждение очистки
                var clearResult = MessageBox.Show($"Вы точно уверены, что хотите удалить все {_exams.Count} экзаменов из базы данных?\n\nЭто действие невозможно отменить!",
                    "Подтверждение очистки", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (clearResult != MessageBoxResult.Yes)
                    return;

                // Очищаем базу данных
                dbhelper.ClearAllExams();

                // Очищаем локальную коллекцию
                _exams.Clear();

                // Обновляем интерфейс
                RefreshExamsData();

                MessageBox.Show($"Буфер успешно очищен. Удалено {_exams.Count} экзаменов.",
                    "Очистка завершена", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке буфера: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion
    }
}
