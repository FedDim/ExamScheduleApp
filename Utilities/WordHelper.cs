using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace ExamScheduleApp.Utilities
{
    public class WordHelper
    {
        private Word.Application _wordApp;
        private Word.Document _doc;
        ObservableCollection<ExamSchedule> _exams;

        public WordHelper(ObservableCollection<ExamSchedule> exams)
        {
            _exams = exams;
        }

        public void CreateAllDocuments(string baseFolderPath)
        {
            try
            {
                string dateFolderName = DateTime.Now.ToString("dd.MM.yyyy");
                string targetFolder = Path.Combine(baseFolderPath, $"Расписания {dateFolderName}");

                if (!Directory.Exists(targetFolder))
                {
                    Directory.CreateDirectory(targetFolder);
                }

                CreateDocumentByDate(Path.Combine(targetFolder, "Расписание промежуточной аттестации (по дате).docx"));
                CreateDocumentByTeachers(Path.Combine(targetFolder, "Расписание промежуточной аттестации (по преподавателям).docx"));
                CreateDocumentBySubjects(Path.Combine(targetFolder, "Расписание промежуточной аттестации (по дисциплинам).docx"));
                CreateDocumentByDepartments(Path.Combine(targetFolder, "Расписание промежуточной аттестации (по отделениям).docx"));

                MessageBox.Show($"Все документы успешно сохранены в папке : \n{targetFolder}", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании документов : {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CreateDocumentByDate(string filePath)
        {
            CreateDocumentInternal("Расписание промежуточной аттестации (по дате)", CreateTableFromExams, filePath);
        }

        private void CreateDocumentByTeachers(string filePath)
        {
            CreateDocumentInternal("Расписание промежуточной аттестации (по преподавателям)", CreateTableFromTeachers, filePath);
        }

        private void CreateDocumentBySubjects(string filePath)
        {
            CreateDocumentInternal("Расписание промежуточной аттестации (по дисциплинам)", CreateTableFromSubjects, filePath);
        }

        private void CreateDocumentByDepartments(string filePath)
        {
            CreateDocumentInternal("Расписание промежуточной аттестации (по отделениям)", CreateTableFromDepartments, filePath);
        }

        private void CreateDocumentInternal(string title, Action tableCreationMethod, string filePath)
        {
            try
            {
                _wordApp = new Word.Application();
                _wordApp.Visible = false;

                _doc = _wordApp.Documents.Add();

                // Настройка параметров страницы
                ConfigurePageSetup();

                // Добавляем нижний колонтитул с номерами страниц
                AddPageNumbers();

                CreateHeader(new List<string> { "Утверждаю: ", "директор СПб ГБПОУ \"АТТ\" ", "_________________Корабельников С.К." });
                CreateDocumentTitle(title);

                // Создаем таблицу
                tableCreationMethod();

                _doc.SaveAs2(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании документа: {ex.Message}\n\nStack Trace: {ex.StackTrace}");
            }
            finally
            {
                Cleanup();
            }
        }

        private void CreateTableFromDepartments()
        {
            try
            {
                // Группируем экзамены по отделениям, включая группы без отделения
                var examsByDepartment = _exams
                    .GroupBy(e => string.IsNullOrEmpty(e.DepartmentName) ? "Без отделения" : e.DepartmentName)
                    .OrderBy(g => g.Key == "Без отделения" ? 1 : 0) // "Без отделения" будет в конце
                    .ThenBy(g => g.Key)
                    .ToList();

                // Подсчитываем общее количество строк
                int totalRows = 1; // заголовок
                foreach (var departmentGroup in examsByDepartment)
                {
                    totalRows++; // строка с отделением

                    // Группируем экзамены внутри отделения по группам
                    var groupsInDepartment = departmentGroup
                        .GroupBy(e => e.GroupName)
                        .OrderBy(g => g.Key)
                        .ToList();

                    foreach (var group in groupsInDepartment)
                    {
                        totalRows++; // строка с группой
                        totalRows += group.Count(); // строки с экзаменами группы
                    }
                }

                // Создаем таблицу с 7 колонками (добавлена Аудитория)
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    totalRows,
                    7, // колонки: Группа, Дата, Время, Дисциплина, Преподаватель, Аудитория, Тип
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                // Убираем границы таблицы
                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                // Устанавливаем повторение заголовков на каждой странице
                table.Rows[1].HeadingFormat = -1;

                // Заголовки таблицы с добавленной колонкой Аудитория
                string[] headers = { "Группа", "Дата", "Время", "Дисциплина", "Преподаватель", "Аудитория", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Cell cell = table.Cell(1, i + 1);
                    cell.Range.Text = headers[i];
                    FormatCell(cell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

                int currentRow = 2;

                // Заполняем таблицу данными
                foreach (var departmentGroup in examsByDepartment)
                {
                    // Добавляем строку с отделением
                    Word.Cell departmentCell = table.Cell(currentRow, 1);
                    departmentCell.Range.Text = departmentGroup.Key;
                    table.Cell(currentRow, 1).Merge(table.Cell(currentRow, 7)); // Объединяем 7 колонок
                    FormatCell(departmentCell, "Times New Roman", 14, true, 14, Word.WdParagraphAlignment.wdAlignParagraphCenter);

                    // Добавляем подчеркивание
                    departmentCell.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;

                    // Белый фон
                    departmentCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    currentRow++;

                    // Группируем экзамены внутри отделения по группам и сортируем по названию группы
                    var groupsInDepartment = departmentGroup
                        .GroupBy(e => e.GroupName)
                        .OrderBy(g => g.Key)
                        .ToList();

                    foreach (var group in groupsInDepartment)
                    {
                        // Добавляем строку с группой
                        Word.Cell groupCell = table.Cell(currentRow, 1);
                        groupCell.Range.Text = group.Key;
                        // Форматируем только ячейку группы (первую колонку)
                        FormatCell(groupCell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);

                        // Остальные ячейки в строке группы оставляем пустыми
                        for (int i = 2; i <= 7; i++)
                        {
                            Word.Cell emptyCell = table.Cell(currentRow, i);
                            emptyCell.Range.Text = "";
                            FormatCell(emptyCell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        }

                        currentRow++;

                        // Сортируем экзамены в группе по дате и времени
                        var sortedExams = group
                            .OrderBy(e => ParseDateForSorting(e.ExamDate) ?? DateTime.MaxValue)
                            .ThenBy(e => ParseTimeForSorting(e.ExamTime) ?? TimeSpan.MaxValue)
                            .ToList();

                        // Добавляем экзамены группы
                        foreach (var exam in sortedExams)
                        {
                            AddDepartmentExamDataToTable(table, currentRow, exam);
                            currentRow++;
                        }
                    }
                }

                // Добавляем отступ после таблицы
                _doc.Range(_doc.Content.End - 1).InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании таблицы по отделениям: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void AddDepartmentExamDataToTable(Word.Table table, int rowNumber, ExamSchedule exam)
        {
            try
            {
                // Формируем строку преподавателя
                string teacherDisplay = GetLastName(exam.Teacher1Name);
                if (!string.IsNullOrEmpty(exam.Teacher2Name))
                {
                    teacherDisplay = $"{GetLastName(exam.Teacher1Name)}/{GetLastName(exam.Teacher2Name)}";
                }

                // Первая ячейка (Группа) оставляем пустой, так как группа уже указана выше
                table.Cell(rowNumber, 1).Range.Text = "";

                // Остальные данные с добавленной колонкой Аудитория
                string[] data = {
                    exam.ExamDate ?? "",
                    exam.ExamTime ?? "",
                    exam.SubjectName ?? "",
                    teacherDisplay,
                    exam.Classroom ?? "", // Используем свойство Classroom
                    exam.ExamType ?? ""
                };

                for (int i = 0; i < data.Length; i++)
                {
                    Word.Cell cell = table.Cell(rowNumber, i + 2); // i+2 потому что первая колонка уже обработана
                    cell.Range.Text = data[i];
                    FormatCell(cell, "Times New Roman", 11, false, 11, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных экзамена в строку {rowNumber}: {ex.Message}");
            }
        }

        private TimeSpan? ParseTimeForSorting(string timeString)
        {
            if (string.IsNullOrEmpty(timeString))
                return null;

            try
            {
                string cleanTime = timeString.Trim();
                if (TimeSpan.TryParse(cleanTime, out TimeSpan time))
                    return time;

                string[] parts = cleanTime.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[0], out int hours) && int.TryParse(parts[1], out int minutes))
                    return new TimeSpan(hours, minutes, 0);
            }
            catch (Exception) { }

            return null;
        }

        private DateTime? ParseDateForSorting(string dateString)
        {
            if (string.IsNullOrEmpty(dateString))
                return null;

            try
            {
                if (DateTime.TryParse(dateString, out DateTime date))
                    return date;
            }
            catch (Exception) { }

            return null;
        }

        private void CreateTableFromSubjects()
        {
            try
            {
                var examsBySubject = _exams
                    .Where(e => !string.IsNullOrEmpty(e.SubjectName))
                    .GroupBy(e => e.SubjectName)
                    .OrderBy(g => g.Key)
                    .ToList();

                int totalRows = 1;
                foreach (var subjectGroup in examsBySubject)
                {
                    totalRows++;
                    totalRows += subjectGroup.Count();
                }

                // Создаем таблицу с 6 колонками (добавлена Аудитория)
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    totalRows,
                    6, // колонки: Дата, Время, Группа, Преподаватель, Аудитория, Тип
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                table.Rows[1].HeadingFormat = -1;

                // Заголовки таблицы с добавленной колонкой Аудитория
                string[] headers = { "Дата", "Время", "Группа", "Преподаватель", "Аудитория", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Cell cell = table.Cell(1, i + 1);
                    cell.Range.Text = headers[i];
                    FormatCell(cell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

                int currentRow = 2;

                foreach (var subjectGroup in examsBySubject)
                {
                    Word.Cell subjectCell = table.Cell(currentRow, 1);
                    subjectCell.Range.Text = subjectGroup.Key;
                    table.Cell(currentRow, 1).Merge(table.Cell(currentRow, 6)); // Объединяем 6 колонок
                    FormatCell(subjectCell, "Times New Roman", 14, true, 14, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                    subjectCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    currentRow++;

                    var sortedExams = subjectGroup
                        .OrderBy(e => ParseDateForSorting(e.ExamDate) ?? DateTime.MaxValue)
                        .ThenBy(e => ParseTimeForSorting(e.ExamTime) ?? TimeSpan.MaxValue)
                        .ToList();

                    foreach (var exam in sortedExams)
                    {
                        AddSubjectExamDataToTable(table, currentRow, exam);
                        currentRow++;
                    }
                }

                _doc.Range(_doc.Content.End - 1).InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании таблицы по дисциплинам: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void AddSubjectExamDataToTable(Word.Table table, int rowNumber, ExamSchedule exam)
        {
            try
            {
                string teacherDisplay = GetLastName(exam.Teacher1Name);
                if (!string.IsNullOrEmpty(exam.Teacher2Name))
                {
                    teacherDisplay = $"{GetLastName(exam.Teacher1Name)}/{GetLastName(exam.Teacher2Name)}";
                }

                // Данные с добавленной колонкой Аудитория
                string[] data = {
                    exam.ExamDate ?? "",
                    exam.ExamTime ?? "",
                    exam.GroupName ?? "",
                    teacherDisplay,
                    exam.Classroom ?? "", // Используем свойство Classroom
                    exam.ExamType ?? ""
                };

                for (int i = 0; i < data.Length; i++)
                {
                    Word.Cell cell = table.Cell(rowNumber, i + 1);
                    cell.Range.Text = data[i];
                    FormatCell(cell, "Times New Roman", 11, false, 11, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных экзамена в строку {rowNumber}: {ex.Message}");
            }
        }

        private void CreateTableFromTeachers()
        {
            try
            {
                var allTeachers = new Dictionary<string, List<ExamSchedule>>();

                foreach (var exam in _exams)
                {
                    if (!string.IsNullOrEmpty(exam.Teacher1Name))
                    {
                        string teacherKey = GetLastName(exam.Teacher1Name);
                        if (!allTeachers.ContainsKey(teacherKey))
                        {
                            allTeachers[teacherKey] = new List<ExamSchedule>();
                        }
                        allTeachers[teacherKey].Add(exam);
                    }

                    if (!string.IsNullOrEmpty(exam.Teacher2Name))
                    {
                        string teacherKey = GetLastName(exam.Teacher2Name);
                        if (!allTeachers.ContainsKey(teacherKey))
                        {
                            allTeachers[teacherKey] = new List<ExamSchedule>();
                        }
                        allTeachers[teacherKey].Add(exam);
                    }
                }

                var sortedTeachers = allTeachers.OrderBy(t => t.Key).ToList();

                int totalRows = 1;
                foreach (var teacher in sortedTeachers)
                {
                    totalRows++;
                    totalRows += teacher.Value.Count;
                }

                // Создаем таблицу с 6 колонками (добавлена Аудитория)
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    totalRows,
                    6, // колонки: Дата, Время, Группа, Дисциплина, Аудитория, Тип
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                table.Rows[1].HeadingFormat = -1;

                // Заголовки таблицы с добавленной колонкой Аудитория
                string[] headers = { "Дата", "Время", "Группа", "Дисциплина", "Аудитория", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Cell cell = table.Cell(1, i + 1);
                    cell.Range.Text = headers[i];
                    FormatCell(cell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

                int currentRow = 2;

                foreach (var teacher in sortedTeachers)
                {
                    Word.Cell teacherCell = table.Cell(currentRow, 1);
                    teacherCell.Range.Text = teacher.Key;
                    table.Cell(currentRow, 1).Merge(table.Cell(currentRow, 6)); // Объединяем 6 колонок
                    FormatCell(teacherCell, "Times New Roman", 14, true, 14, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                    teacherCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    currentRow++;

                    var sortedExams = teacher.Value
                        .OrderBy(e => ParseDateForSorting(e.ExamDate) ?? DateTime.MaxValue)
                        .ThenBy(e => ParseTimeForSorting(e.ExamTime) ?? TimeSpan.MaxValue)
                        .ToList();

                    foreach (var exam in sortedExams)
                    {
                        AddTeacherExamDataToTable(table, currentRow, exam);
                        currentRow++;
                    }
                }

                _doc.Range(_doc.Content.End - 1).InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании таблицы по преподавателям: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void AddTeacherExamDataToTable(Word.Table table, int rowNumber, ExamSchedule exam)
        {
            try
            {
                // Данные с добавленной колонкой Аудитория
                string[] data = {
                    exam.ExamDate ?? "",
                    exam.ExamTime ?? "",
                    exam.GroupName ?? "",
                    exam.SubjectName ?? "",
                    exam.Classroom ?? "", // Используем свойство Classroom
                    exam.ExamType ?? ""
                };

                for (int i = 0; i < data.Length; i++)
                {
                    Word.Cell cell = table.Cell(rowNumber, i + 1);
                    cell.Range.Text = data[i];
                    FormatCell(cell, "Times New Roman", 11, false, 11, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных экзамена в строку {rowNumber}: {ex.Message}");
            }
        }

        private void CreateTableFromExams()
        {
            try
            {
                var examsByDate = _exams
                    .Where(e => !string.IsNullOrEmpty(e.ExamDate))
                    .GroupBy(e => e.ExamDate)
                    .OrderBy(g => ParseDateForSorting(g.Key) ?? DateTime.MaxValue)
                    .ToList();

                int totalRows = 1;
                foreach (var dateGroup in examsByDate)
                {
                    totalRows++;
                    totalRows += dateGroup.Count();
                }

                // Создаем таблицу с 6 колонками (добавлена Аудитория)
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    totalRows,
                    6, // колонки: Время, Группа, Дисциплина, Преподаватель, Аудитория, Тип
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                table.Rows[1].HeadingFormat = -1;

                // Заголовки таблицы с добавленной колонкой Аудитория
                string[] headers = { "Время", "Группа", "Дисциплина", "Преподаватель", "Аудитория", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Cell cell = table.Cell(1, i + 1);
                    cell.Range.Text = headers[i];
                    FormatCell(cell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

                int currentRow = 2;

                foreach (var dateGroup in examsByDate)
                {
                    string dayOfWeek = CapitalizeFirstLetter(GetDayOfWeekFromDate(dateGroup.Key));
                    string dateDisplay = $"{dateGroup.Key} {dayOfWeek}";

                    Word.Cell dateCell = table.Cell(currentRow, 1);
                    dateCell.Range.Text = dateDisplay;
                    table.Cell(currentRow, 1).Merge(table.Cell(currentRow, 6)); // Объединяем 6 колонок
                    FormatCell(dateCell, "Times New Roman", 14, true, 14, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                    dateCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    currentRow++;

                    var sortedExams = dateGroup
                        .OrderBy(e => ParseTimeForSorting(e.ExamTime) ?? TimeSpan.MaxValue)
                        .ToList();

                    foreach (var exam in sortedExams)
                    {
                        AddExamDataToTable(table, currentRow, exam);
                        currentRow++;
                    }
                }

                _doc.Range(_doc.Content.End - 1).InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании таблицы из данных: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void AddExamDataToTable(Word.Table table, int rowNumber, ExamSchedule exam)
        {
            try
            {
                string teacherDisplay = GetLastName(exam.Teacher1Name);
                if (!string.IsNullOrEmpty(exam.Teacher2Name))
                {
                    teacherDisplay = $"{GetLastName(exam.Teacher1Name)}/{GetLastName(exam.Teacher2Name)}";
                }

                // Данные с добавленной колонкой Аудитория
                string[] data = {
                    exam.ExamTime ?? "",
                    exam.GroupName ?? "",
                    exam.SubjectName ?? "",
                    teacherDisplay,
                    exam.Classroom ?? "", // Используем свойство Classroom
                    exam.ExamType ?? ""
                };

                for (int i = 0; i < data.Length; i++)
                {
                    Word.Cell cell = table.Cell(rowNumber, i + 1);
                    cell.Range.Text = data[i];
                    FormatCell(cell, "Times New Roman", 11, false, 11, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных экзамена в строку {rowNumber}: {ex.Message}");
            }
        }

        private string GetLastName(string fullName)
        {
            if (string.IsNullOrEmpty(fullName))
                return "";

            string[] nameParts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return nameParts.Length > 0 ? nameParts[0] : fullName;
        }

        private void AddPageNumbers()
        {
            try
            {
                object missing = Type.Missing;

                foreach (Word.Section section in _doc.Sections)
                {
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    footer.Range.Delete();
                    footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    footer.Range.Font.Name = "Times New Roman";
                    footer.Range.Font.Size = 10;

                    footer.Range.Select();
                    Word.Selection selection = _wordApp.Selection;

                    selection.TypeText("Страница ");
                    selection.Fields.Add(selection.Range, Word.WdFieldType.wdFieldPage);
                    selection.TypeText(" из ");
                    selection.Fields.Add(selection.Range, Word.WdFieldType.wdFieldNumPages);

                    selection.Fields.Update();
                }

                _wordApp.Selection.Collapse();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении номеров страниц: {ex.Message}");
            }
        }

        private void ConfigurePageSetup()
        {
            try
            {
                Word.PageSetup pageSetup = _doc.PageSetup;
                pageSetup.TopMargin = _wordApp.CentimetersToPoints(2.0f);
                pageSetup.BottomMargin = _wordApp.CentimetersToPoints(2.0f);
                pageSetup.LeftMargin = _wordApp.CentimetersToPoints(2.1f);
                pageSetup.RightMargin = _wordApp.CentimetersToPoints(1.0f);
                pageSetup.Gutter = _wordApp.CentimetersToPoints(0f);
                pageSetup.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка в настройке страницы: {ex.Message}");
            }
        }

        private void CreateHeader(List<string> lines)
        {
            if (lines == null || lines.Count == 0)
                return;

            for (int i = 0; i < lines.Count; i++)
            {
                Word.Paragraph paragraph = _doc.Content.Paragraphs.Add();

                paragraph.Range.Text = "\t" + lines[i];

                SafeSetFont(paragraph.Range.Font, "Times New Roman", 12, 1);

                paragraph.Format.SpaceAfter = 0;
                paragraph.Format.SpaceBefore = (i == 0) ? 12 : 0;
                paragraph.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                paragraph.Format.TabStops.Add(_wordApp.CentimetersToPoints(9.5f));

                paragraph.Range.InsertParagraphAfter();
            }
        }

        private void CreateDocumentTitle(string titleText)
        {
            Word.Paragraph titleParagraph = _doc.Content.Paragraphs.Add();

            titleParagraph.Range.Text = titleText;

            SafeSetFont(titleParagraph.Range.Font, "Times New Roman", 18, 1);

            titleParagraph.Format.SpaceAfter = 0;
            titleParagraph.Format.SpaceBefore = 18;
            titleParagraph.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            titleParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            titleParagraph.Range.InsertParagraphAfter();
        }

        private string GetDayOfWeekFromDate(string dateString)
        {
            try
            {
                if (DateTime.TryParse(dateString, out DateTime date))
                {
                    return date.ToString("dddd", new System.Globalization.CultureInfo("ru-RU"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при преобразовании даты: {ex.Message}");
            }

            return "Неизвестный день";
        }

        private string CapitalizeFirstLetter(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return char.ToUpper(text[0]) + text.Substring(1).ToLower();
        }

        private void FormatCell(Word.Cell cell, string fontName, float fontSize, bool isBold, float spaceBefore, Word.WdParagraphAlignment alignment)
        {
            try
            {
                cell.Range.Font.Name = fontName;
                cell.Range.Font.Size = fontSize;
                cell.Range.Font.Bold = isBold ? 1 : 0;
                cell.Range.ParagraphFormat.SpaceAfter = 0;
                cell.Range.ParagraphFormat.SpaceBefore = spaceBefore;
                cell.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                cell.Range.ParagraphFormat.Alignment = alignment;
                cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка форматирования ячейки: {ex.Message}");
            }
        }

        private void SafeSetFont(Word.Font font, string fontName, float fontSize, int bold)
        {
            try
            {
                font.Name = fontName;
                font.Size = fontSize;
                font.Bold = bold;
            }
            catch (COMException)
            {
                font.Name = "Arial";
                font.Size = fontSize;
                font.Bold = bold;
            }
        }

        private void Cleanup()
        {
            try
            {
                if (_doc != null)
                {
                    _doc.Close();
                    Marshal.ReleaseComObject(_doc);
                    _doc = null;
                }
                if (_wordApp != null)
                {
                    _wordApp.Quit();
                    Marshal.ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при очистке ресурсов: {ex.Message}");
            }
        }
    }
}