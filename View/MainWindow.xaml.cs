using ExamScheduleApp.Model;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
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
        private string _dataFolder = "Data";
        private ObservableCollection<ExamSchedule> _exams = new ObservableCollection<ExamSchedule>();

        public string ReceivedData { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            LoadItemsFromFile(GetFilePathFromData("teachers.txt"), cbTeachers);
            LoadItemsFromFile(GetFilePathFromData("teachers.txt"), SecondTeacherCB);
            LoadItemsFromFile(GetFilePathFromData("disciplines.txt"), cbSubjects);
            LoadItemsFromFile(GetFilePathFromData("groups.txt"), GroupComboBox);
        }

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

        // Добавление преподавателя
        private void AddTeacher_Click(object sender, RoutedEventArgs e)
        {
            Window addWindow = new AddWindow("фамилию преподавателя");

            bool? result = addWindow.ShowDialog();

            string filePath = GetFilePathFromData("teachers.txt");
            string newLine = AddWindow.data.ToUpper();
            try
            {
                List<string> lines = new List<string>();

                // Чтение существующих строк, если файл существует
                if (File.Exists(filePath))
                {
                    lines = File.ReadAllLines(filePath).ToList();
                }

                // Добавление новой строки
                if (!string.IsNullOrWhiteSpace(newLine))
                {
                    lines.Add(newLine.Trim());
                }

                // Сортировка строк в алфавитном порядке
                lines.Sort();

                //Удаление дубликатов
                var uniqueLines = new HashSet<string>(lines);

                // Запись отсортированных строк обратно в файл
                File.WriteAllLines(filePath, uniqueLines);
                LoadItemsFromFile(filePath, cbTeachers);
                LoadItemsFromFile(filePath, SecondTeacherCB);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с файлом: {ex.Message}", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }

        }

        // Добавление дисциплины
        private void AddSubject_Click(object sender, RoutedEventArgs e)
        {
            Window addWindow = new AddWindow("дисциплину");

            bool? result = addWindow.ShowDialog();


            string filePath = GetFilePathFromData("disciplines.txt");
            string newLine = AddWindow.data;
            try
            {
                List<string> lines = new List<string>();

                // Чтение существующих строк, если файл существует
                if (File.Exists(filePath))
                {
                    lines = File.ReadAllLines(filePath).ToList();
                }

                // Добавление новой строки
                if (!string.IsNullOrWhiteSpace(newLine))
                {
                    lines.Add(newLine.Trim());
                }

                // Сортировка строк в алфавитном порядке
                lines.Sort();

                //Удаление дубликатов
                var uniqueLines = new HashSet<string>(lines);

                // Запись отсортированных строк обратно в файл
                File.WriteAllLines(filePath, uniqueLines);
                LoadItemsFromFile(filePath, cbSubjects);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с файлом: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private void UpdateStatus(string message)
        {
            tbStatus.Text = $"{DateTime.Now:HH:mm:ss}: {message}";
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

            ExamsDataGrid.ItemsSource = _exams;

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

        private void AddGroup_Click(object sender, RoutedEventArgs e)
        {
            Window addWindow = new AddWindow("группу");

            bool? result = addWindow.ShowDialog();


            string filePath = GetFilePathFromData("groups.txt");
            string newLine = AddWindow.data;
            try
            {
                List<string> lines = new List<string>();

                // Чтение существующих строк, если файл существует
                if (File.Exists(filePath))
                {
                    lines = File.ReadAllLines(filePath).ToList();
                }

                // Добавление новой строки
                if (!string.IsNullOrWhiteSpace(newLine))
                {
                    lines.Add(newLine.Trim());
                }

                // Сортировка строк в алфавитном порядке
                lines.Sort();

                //Удаление дубликатов
                var uniqueLines = new HashSet<string>(lines);

                // Запись отсортированных строк обратно в файл
                File.WriteAllLines(filePath, uniqueLines);
                LoadItemsFromFile(filePath, GroupComboBox);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при работе с файлом: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private string GetFilePathFromData(string file)
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
