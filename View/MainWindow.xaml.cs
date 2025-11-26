using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SQLite;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

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
        private ObservableCollection<ExamSchedule> _exams = new ObservableCollection<ExamSchedule>();
        public SimpleDatabaseHelper dbhelper = new SimpleDatabaseHelper();

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                // Диагностика базы данных
                dbhelper.CheckDatabaseStructure();

                // Очищаем проблемные данные (для SimpleDatabaseHelper.cs)
                dbhelper.CleanProblematicData();

                // Очищаем все экзамены при запуске
                dbhelper.ClearAllExams();

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

                // DataGrid с экзаменами
                var examList = dbhelper.GetExamSchedule();
                _exams.Clear();
                foreach (var exam in examList)
                {
                    _exams.Add(exam);
                }
                ExamsDataGrid.Items.Refresh();

                UpdateStatus("Данные обновлены");

                // Убрать после наладки
                //MessageBox.Show($"Загружено: {_teachers.Count} преподавателей, {_subjects.Count} дисциплин, {_groups.Count} групп, {_exams.Count} экзаменов");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных из базы: {ex.Message}");
                UpdateStatus($"Ошибка загрузки: {ex.Message}");
            }
        }
        private void AddExam(object sender, RoutedEventArgs e)
        {
            if (cbTeachers.SelectedItem == null || SecondTeacherCB.SelectedItem == null ||
                cbSubjects.SelectedItem == null || GroupComboBox.SelectedItem == null)
            {
                MessageBox.Show("Выберите всех преподавателей, дисциплину и группу!");
                return;
            }

            // Получаем объекты
            var teacher1 = (Teacher)cbTeachers.SelectedItem;
            var teacher2 = (Teacher)SecondTeacherCB.SelectedItem;
            var subject = (Subject)cbSubjects.SelectedItem;
            var group = (Group)GroupComboBox.SelectedItem;

            string surname = teacher1.Name;
            string secondSurname = teacher2.Name;

            // Дата, время и тип вводятся вручную
            DateTime? date = dpExamDate.SelectedDate;
            string dateString = date?.ToString("dd.MM.yyyy") ?? string.Empty;
            string time = TimeComboBox.SelectedItem != null ? TimeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;
            string type = TypeComboBox.SelectedItem != null ? TypeComboBox.SelectedItem.ToString().Replace("System.Windows.Controls.ComboBoxItem: ", "") : string.Empty;

            string subjectName = subject.Name;
            string groupName = group.Name;
            string cabinet = txtClassroom.Text;

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
                    Teacher2Id = teacher2.Id,
                    SubjectId = subject.Id,
                    GroupId = group.Id,
                    Classroom = cabinet,
                    Teacher1Name = surname,
                    Teacher2Name = secondSurname,
                    SubjectName = subjectName,
                    GroupName = groupName,
                    // Дата, время и тип сохраняются только в объекте, не в БД
                    ExamDate = dateString,
                    ExamTime = time,
                    ExamType = type
                };

                // Добавление в базу данных (без даты, времени и типа)
                dbhelper.AddExam(exam);

                // Получаем ID добавленной записи
                var addedExams = dbhelper.GetExamSchedule();
                if (addedExams.Count > 0)
                {
                    exam.Id = addedExams[addedExams.Count - 1].Id;
                }

                // Обновление интерфейса
                _exams.Add(exam);
                ExamsDataGrid.Items.Refresh();

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
            if (string.IsNullOrEmpty(secondSurname)) checkResult += "Не выбрана Фамилия второго преподавателя\n";
            if (string.IsNullOrEmpty(date)) checkResult += "Не выбрана дата проведения\n";
            if (string.IsNullOrEmpty(subject)) checkResult += "Не выбрана Дисциплина\n";
            if (string.IsNullOrEmpty(time)) checkResult += "Не заполнено Время\n";
            if (string.IsNullOrEmpty(cabinet)) checkResult += "Не заполнен Номер Кабинета\n";
            if (string.IsNullOrEmpty(type)) checkResult += "Не выбран Тип\n";

            if (!checkResult.Equals(string.Empty))
                checkResult = "Не все данные заполнены: \n" + checkResult;

            return checkResult;
        }

        private void UpdateStatus(string message)
        {
            tbStatus.Text = $"{DateTime.Now:HH:mm:ss}: {message}";
        }

        private void GenerateWord_Click(object sender, RoutedEventArgs e)
        {
            // Создаем экземпляр Word
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();

            try
            {
                // Настройка полей документа (в пунктах) ~3 см
                doc.PageSetup.LeftMargin = 85;
                doc.PageSetup.RightMargin = 85;
                doc.PageSetup.TopMargin = 85;
                doc.PageSetup.BottomMargin = 85;

                // Создание шапки документа
                CreateHeader(doc);

                // Добавление таблицы с данными
                CreateExamTable(doc);

                string documentsPath = GetDocumentsFolderPath();

                string filePath = ShowSaveFileDialog(documentsPath);

                if (!string.IsNullOrEmpty(filePath))
                {
                    doc.SaveAs2(filePath);
                    MessageBox.Show($"Файл успешно сохранён по пути : {filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        private void CreateHeader(Word.Document doc)
        {
            // Первый параграф - "Утверждаю:"
            Word.Paragraph p1 = doc.Content.Paragraphs.Add();
            p1.Range.Text = "Утверждаю:";
            p1.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            p1.Range.Font.Name = "Times New Roman";
            p1.Range.Font.Size = 12;
            p1.Range.InsertParagraphAfter();

            // Второй параграф - должность
            Word.Paragraph p2 = doc.Content.Paragraphs.Add();
            p2.Range.Text = "директор СПб ГБПОУ \"АТТ\"";
            p2.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            p2.Range.Font.Name = "Times New Roman";
            p2.Range.Font.Size = 12;
            p2.Range.InsertParagraphAfter();

            // Третий параграф - подпись
            Word.Paragraph p3 = doc.Content.Paragraphs.Add();
            p3.Range.Text = "___________________ Корабельников С.К.";
            p3.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            p3.Range.Font.Name = "Times New Roman";
            p3.Range.Font.Size = 12;
            p3.Format.SpaceAfter = 24; // Больший отступ после подписи
            p3.Range.InsertParagraphAfter();

            // Четвертый параграф - заголовок
            Word.Paragraph p4 = doc.Content.Paragraphs.Add();
            p4.Range.Text = "Расписание промежуточной аттестации (по дате)";
            p4.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            p4.Range.Font.Name = "Times New Roman";
            p4.Range.Font.Size = 18;
            p4.Range.Font.Bold = 1;
            p4.Format.SpaceAfter = 18;
            p4.Range.InsertParagraphAfter();
        }

        private void CreateExamTable(Word.Document doc)
        {
            int rowCount = ExamsDataGrid.Items.Count;
            int columnCount = ExamsDataGrid.Columns.Count;

            // Добавляем строку для заголовков
            Word.Table wordTable = doc.Tables.Add(
                doc.Range(doc.Content.End - 1),
                rowCount + 1,
                columnCount,
                Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                Word.WdAutoFitBehavior.wdAutoFitWindow
            );

            wordTable.Range.Font.Name = "Times New Roman";
            wordTable.Range.Font.Size = 12;
            wordTable.Range.Font.Bold = 0; // 0 - не жирный, 1 - жирный
            wordTable.Range.Font.Italic = 0; // 0 - не курсив, 1 - курсив
            wordTable.Range.Font.Color = Word.WdColor.wdColorBlack;

            wordTable.Borders.Enable = 0;
            wordTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            wordTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

            // Заполняем данные
            for (int row = 0; row < rowCount; row++)
            {
                var item = ExamsDataGrid.Items[row];
                for (int col = 0; col < columnCount; col++)
                {
                    var cellValue = GetCellValue(item, ExamsDataGrid.Columns[col]);
                    wordTable.Cell(row + 2, col + 1).Range.Text = cellValue?.ToString() ?? "";
                }
            }

            // Форматирование таблицы
            wordTable.Range.Font.Size = 10;
        }

        private object GetCellValue(object item, DataGridColumn dataGridColumn)
        {
            if (dataGridColumn is DataGridBoundColumn boundColumn)
            {
                var binding = boundColumn.Binding as System.Windows.Data.Binding;
                if (binding?.Path != null)
                {
                    string propertyName = binding.Path.Path;
                    PropertyInfo prop = item.GetType().GetProperty(propertyName);
                    return prop?.GetValue(item);
                }
            }
            return null;
        }

        private string GetDocumentsFolderPath()
        {
            string documentsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents");

            if (!Directory.Exists(documentsPath))
            {
                Directory.CreateDirectory(documentsPath);
            }

            return documentsPath;
        }

        private string ShowSaveFileDialog(string initalDirectory)
        {
            var saveFileDialog = new SaveFileDialog
            {
                InitialDirectory = initalDirectory,
                FileName = $"Экзамены_{DateTime.Now:dd.MM.yyyy}",
                Filter = "Word Документы (*.docx)|*.docx|Все файлы (*.*)|*.*",
                DefaultExt = ".docx",
                AddExtension = true,
                Title = "Сохранить документ Word"
            };

            bool? result = saveFileDialog.ShowDialog();

            return result == true ? saveFileDialog.FileName : null;
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

        private void DeleteExam_Click(object sender, RoutedEventArgs e)
        {
            if (ExamsDataGrid.SelectedItem == null)
            {
                MessageBox.Show("Выберите экзамен для удаления!");
                return;
            }

            var selectedExam = (ExamSchedule)ExamsDataGrid.SelectedItem;

            var result = MessageBox.Show(
                $"Вы действительно хотите удалить экзамен?\n" +
                $"Преподаватели: {selectedExam.Teacher1Name}, {selectedExam.Teacher2Name}\n" +
                $"Дисциплина: {selectedExam.SubjectName}\n" +
                $"Группа: {selectedExam.GroupName}",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // Удаляем из базы данных
                    dbhelper.DeleteExam(selectedExam.Id);

                    // Удаляем из коллекции
                    _exams.Remove(selectedExam);
                    ExamsDataGrid.Items.Refresh();

                    MessageBox.Show("Экзамен удален!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении экзамена: {ex.Message}");
                }
            }
        }

        private void EditTeachers_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditTeachersWindow();
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditSubjects_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditSubjectsWindow();
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditGroups_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditGroupsWindow();
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }
    }
}
