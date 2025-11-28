using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using ExamScheduleApp.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ExamScheduleApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private List<Teacher> _teachers;
        private List<Subject> _subjects;
        private string _dataFolder = "Data";
        private ObservableCollection<ExamSchedule> _exams = new ObservableCollection<ExamSchedule>();
        private string _bufferFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Buffer");
        private string _bufferFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Buffer/буфер.txt");

        public string ReceivedData { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            InitializeBuffer();

            LoadItemsFromFile(FileHelper.GetFilePathFromData("teachers.txt"), cbTeachers);
            LoadItemsFromFile(FileHelper.GetFilePathFromData("teachers.txt"), SecondTeacherCB);
            LoadItemsFromFile(FileHelper.GetFilePathFromData("disciplines.txt"), cbSubjects);
            LoadItemsFromFile(FileHelper.GetFilePathFromData("groups.txt"), GroupComboBox);
        }

        #region Буфер
        private void InitializeBuffer()
        {

            if (!Directory.Exists(_bufferFolder))
            {
                Directory.CreateDirectory(_bufferFolder);
            }

            if (File.Exists(_bufferFile))
            {
                var result = MessageBox.Show("Обнаружен файл буфера. Хотите загрузить данные из буфера?", "Загрузка Буфера",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes) LoadFromBuffer();
                else
                {
                    RenameCurrentBuffer();
                    _exams.Clear();
                }

                CleanOldBufferFiles();
            }
        }

        private void LoadFromBuffer()
        {
            if (!File.Exists(_bufferFile))
            {
                MessageBox.Show("Файл буфера не найден");
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(_bufferFile, Encoding.UTF8);
                int errorCount = 0;
                bool showErrorLine = true;

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    string[] parts = line.Split('|');

                    if (parts.Length == 8)
                    {
                        for (int j = 0; j < parts.Length; j++)
                        {
                            parts[j] = parts[j].Trim();
                        }

                        ExamSchedule exam = new ExamSchedule
                        (
                            parts[0], parts[1], parts[2], parts[3],
                            parts[4], parts[5], parts[6], parts[7]
                        );
                        _exams.Add(exam);
                    }
                    else
                    {
                        errorCount++;

                        if (showErrorLine)
                        {
                            MessageBox.Show($"Строка {i + 1} имеет неверный формат и будет пропущена :\n{line}", "Ошибка Формата",
                                MessageBoxButton.OK, MessageBoxImage.Warning);

                            if (errorCount == 1 || errorCount % 10 == 0)
                            {
                                var result = MessageBox.Show("Продолжать уведомлять о каждой неверной строке?", "Система оповещения",
                                    MessageBoxButton.YesNo, MessageBoxImage.Question);

                                if (result == MessageBoxResult.No) showErrorLine = false;
                            }
                        }
                    }
                }

                if (errorCount > 0) MessageBox.Show($"Загрузка завершена. Пропущено {errorCount} строк с ошибками.");
                else MessageBox.Show("Данные успешно загружены из буфера.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке буфера: {ex.Message}");
            }
        }

        private void RenameCurrentBuffer()
        {
            if (File.Exists(_bufferFile))
            {
                try
                {
                    string dateString = DateTime.Now.ToString("dd.MM.yyyy_HH-mm-ss");
                    string newFileName = Path.Combine(_bufferFolder, $"/буфер_{dateString}.txt");

                    int counter = 1;
                    while (File.Exists(newFileName))
                    {
                        newFileName = Path.Combine(_bufferFolder, $"/буфер_{dateString}_{counter}.txt");
                        counter++;
                    }

                    File.Move(_bufferFile, newFileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при переименовании буфера: {ex.Message}");
                }
            }
        }

        private void CleanOldBufferFiles()
        {
            try
            {
                var bufferFiles = Directory.GetFiles(_bufferFolder, "буфер_*.txt")
                    .Select(file => new FileInfo(file))
                    .OrderByDescending(file => file.CreationTime)
                    .ToList();

                for (int i = 5; i < bufferFiles.Count; i++)
                {
                    bufferFiles[i].Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке старых буферов: {ex.Message}");
            }
        }

        private void SaveToBuffer()
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(_bufferFile, false, Encoding.UTF8))
                {
                    foreach (var exam in _exams)
                    {
                        string line = $"{exam.FirstTeacher} | {exam.SecondTeacher} | {exam.ExamDate} | {exam.Subject} | " +
                            $"{exam.Group} | {exam.ExamTime} | {exam.Classroom} | {exam.ExamType}";
                        writer.WriteLine(line);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении буфера: {ex.Message}");
            }
        }

        private void CreateBackupBuffer()
        {
            if (File.Exists(_bufferFolder))
            {
                try
                {
                    string dateString = DateTime.Now.ToString("dd.MM.yyyy_HH-mm-ss");
                    string newFileName = Path.Combine(_bufferFolder, $"/буфер_{dateString}.txt");

                    int counter = 1;
                    while (File.Exists(newFileName))
                    {
                        newFileName = Path.Combine(_bufferFolder, $"/буфер_{dateString}_{counter}.txt");
                        counter++;
                    }

                    File.Copy(_bufferFolder, newFileName);
                    File.Delete(_bufferFolder);
                    MessageBox.Show($"Буфер сохранен как: {Path.GetFileName(newFileName)}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании резервной копии буфера: {ex.Message}");
                }
            }
        }
        #endregion

        private void InitializeData()
        {
            _teachers = new List<Teacher>();
            _subjects = new List<Subject>();

            // Создаем папку для данных если не существует
            if (!Directory.Exists(_dataFolder))
            {
                Directory.CreateDirectory(_dataFolder);
            }
        }

        private void LoadItemsFromFile(string filePath, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            try
            {

                // Проверяем существование файла
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Файл {filePath} не найден");
                    return;
                }


                // Читаем все строки из файла
                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);

                // Добавляем каждую строку в ComboBox
                foreach (string line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        comboBox.Items.Add(line.Trim());
                    }
                }

                // Устанавливаем первый элемент как выбранный (опционально)
                if (comboBox.Items.Count > 0)
                    comboBox.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла: {ex.Message}");
            }
        }

        private void ShowScheduleTable_Click(object sender, RoutedEventArgs e)
        {
            var scheduleTableWindow = new ScheduleTableWindow(_exams);
            scheduleTableWindow.Owner = this;
            scheduleTableWindow.ShowDialog();
        }

        private void AddExam(object sender, RoutedEventArgs e)
        {
            string surname = cbTeachers.SelectedItem.ToString();
            string secondSurname = SecondTeacherCB.SelectedItem.ToString();
            DateTime? date = dpExamDate.SelectedDate;
            string dateString = date?.ToString("dd.MM.yyyy") ?? string.Empty;
            string subject = cbSubjects.SelectedItem.ToString();
            string group = GroupComboBox.SelectedItem.ToString();
            string time = TimeComboBox.SelectedItem != null ? TimeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;
            string cabinet = txtClassroom.Text;
            string type = TypeComboBox.SelectedItem != null ? TypeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;

            string checkResult = CheckEnteredFileds(surname, secondSurname, dateString, subject, group, time, cabinet, type);

            if (!checkResult.Equals(string.Empty))
            {
                MessageBox.Show(checkResult);
                return;
            }

            ExamSchedule exam = new ExamSchedule(surname, secondSurname, dateString, subject, group, time, cabinet, type);

            _exams.Add(exam);

            SaveToBuffer();

        }

        private string CheckEnteredFileds(string surname, string secondSurname, string date, string subject, string group, string time, string cabinet, string type)
        {
            string checkResult = string.Empty;

            if (surname.Equals(string.Empty)) checkResult += "Не выбрана Фамилия первого преподователя\n";
            if (secondSurname.Equals(string.Empty)) checkResult += "Не выбрана Фамилия второго преподователя\n";
            if (date.Equals(string.Empty)) checkResult += "Не выбрана дата проведения\n";
            if (subject.Equals(string.Empty)) checkResult += "Не выбрана Дисциплина\n";
            if (time.Equals(string.Empty)) checkResult += "Не заполнено Время\n";
            if (cabinet.Equals(string.Empty)) checkResult += "Не заполнен Номер Кaбинета\n";
            if (type.Equals(string.Empty)) checkResult += "Не выбран Тип\n";

            if (!checkResult.Equals(string.Empty)) checkResult = "Не все данные заполены : \n" + checkResult;

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
                    LoadItemsFromFile(FileHelper.GetFilePathFromData("teachers.txt"), cbTeachers);
                    LoadItemsFromFile(FileHelper.GetFilePathFromData("teachers.txt"), SecondTeacherCB);
                    break;
                case DataType.SUBJECT:
                    LoadItemsFromFile(FileHelper.GetFilePathFromData("disciplines.txt"), cbSubjects);
                    break;
                case DataType.GROUP:
                    LoadItemsFromFile(FileHelper.GetFilePathFromData("groups.txt"), GroupComboBox);
                    break;
            }
        }
        #endregion

        private void LoadFromBufferButton_Click(object sender, RoutedEventArgs e)
        {
            if (_exams.Count > 0)
            {
                var result = MessageBox.Show("Текущие данные будут потеряны. Продолжить?",
                    "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.No)
                    return;
            }

            if (_exams.Count > 0 && File.Exists(_bufferFile))
            {
                CreateBackupBuffer();
            }

            _exams.Clear();
            LoadFromBuffer();
        }

        private void ClearBufferButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите очистить буфер? Текущие данные будут сохранены в архив.",
    "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                if (File.Exists(_bufferFile) && _exams.Count > 0)
                {
                    CreateBackupBuffer();
                }
                _exams.Clear();
                MessageBox.Show("Буфер очищен и сохранен в архив.");
            }
        }

        private void NewTableButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Начать новую таблицу? Текущие данные будут сохранены в архив.",
    "Новая таблица", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                if (File.Exists(_bufferFile) && _exams.Count > 0)
                {
                    CreateBackupBuffer();
                }
                _exams.Clear();
                MessageBox.Show("Начата новая таблица.");
            }
        }
    }
}
