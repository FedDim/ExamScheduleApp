using ExamScheduleApp.Model;
using Microsoft.Win32;
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

        public void CreateDocument()
        {
            try
            {
                _wordApp = new Word.Application();
                _wordApp.Visible = true;

                _doc = _wordApp.Documents.Add();

                // Настройка параметров страницы
                ConfigurePageSetup();

                // Добавляем нижний колонтитул с номерами страниц
                AddPageNumbers();

                CreateHeader(new List<string> { "Утверждаю: ", "директор СПб ГБПОУ \"АТТ\" ", "_________________Корабельников С.К." });
                CreateDocumentTitle("Расписание промежуточной аттестации (по дате)");

                // Создаем таблицу из реальных данных или тестовую
                if (_exams != null && _exams.Count > 0)
                {
                    CreateTableFromExams();
                }
                else
                {
                    CreateSimpleTable(); // для тестовых данных
                }

                string documentsPath = GetDocumentsFolderPath();
                string filePath = ShowSaveFileDialog(documentsPath);

                if (!string.IsNullOrEmpty(filePath))
                {
                    _doc.SaveAs2(filePath);
                    MessageBox.Show($"Файл успешно сохранён по пути: {filePath}");
                }
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

        private void AddPageNumbers()
        {
            try
            {
                // Добавляем нижний колонтитул для всех разделов
                foreach (Word.Section section in _doc.Sections)
                {
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    // Очищаем существующий контент
                    footer.Range.Delete();

                    // Выравниваем по правому краю
                    footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    // Добавляем текст "Страница X из Y"
                    footer.Range.Text = "Страница ";
                    Word.Field pageField = footer.Range.Fields.Add(
                        footer.Range,
                        Word.WdFieldType.wdFieldPage,
                        Text: "",
                        PreserveFormatting: true
                    );
                    footer.Range.Text = " из ";
                    Word.Field numPagesField = footer.Range.Fields.Add(
                        footer.Range,
                        Word.WdFieldType.wdFieldNumPages,
                        Text: "",
                        PreserveFormatting: true
                    );

                    // Форматируем шрифт
                    footer.Range.Font.Name = "Times New Roman";
                    footer.Range.Font.Size = 10;

                    // Обновляем поля
                    pageField.Update();
                    numPagesField.Update();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении номеров страниц: {ex.Message}");
            }
        }

        private void CreateTableFromExams()
        {
            try
            {
                // Группируем экзамены по дате и сортируем по дате
                var examsByDate = _exams
                    .Where(e => !string.IsNullOrEmpty(e.ExamDate))
                    .GroupBy(e => e.ExamDate)
                    .OrderBy(g => DateTime.Parse(g.Key)) // Сортируем по дате
                    .ToList();

                // Подсчитываем общее количество строк
                int totalRows = 1; // заголовок
                foreach (var dateGroup in examsByDate)
                {
                    totalRows++; // строка с датой
                    totalRows += dateGroup.Count(); // строки с экзаменами
                }

                // Создаем таблицу
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    totalRows,
                    5, // колонки: Время, Группа, Дисциплина, Преподаватель, Тип
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                // Убираем границы таблицы
                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                // Устанавливаем повторение заголовков на каждой странице
                table.Rows[1].HeadingFormat = -1; // -1 соответствует true в Word Interop

                // Заголовки таблицы
                string[] headers = { "Время", "Группа", "Дисциплина", "Преподаватель", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    Word.Cell cell = table.Cell(1, i + 1);
                    cell.Range.Text = headers[i];
                    FormatCell(cell, "Times New Roman", 12, true, 12, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

                int currentRow = 2;

                // Заполняем таблицу данными
                foreach (var dateGroup in examsByDate)
                {
                    // Получаем день недели из даты с заглавной буквы
                    string dayOfWeek = CapitalizeFirstLetter(GetDayOfWeekFromDate(dateGroup.Key));
                    string dateDisplay = $"{dateGroup.Key} {dayOfWeek}";

                    // Добавляем строку с датой
                    Word.Cell dateCell = table.Cell(currentRow, 1);
                    dateCell.Range.Text = dateDisplay;
                    table.Cell(currentRow, 1).Merge(table.Cell(currentRow, 5));
                    FormatCell(dateCell, "Times New Roman", 14, true, 14, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                    // Белый фон
                    dateCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                    currentRow++;

                    // Сортируем экзамены в группе по времени и добавляем их
                    var sortedExams = dateGroup
                        .OrderBy(e => e.ExamTime)
                        .ToList();

                    foreach (var exam in sortedExams)
                    {
                        AddExamDataToTable(table, currentRow, exam);
                        currentRow++;
                    }
                }

                // Добавляем отступ после таблицы
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
                // Формируем строку преподавателя - только фамилии
                string teacherDisplay = GetLastName(exam.FirstTeacher);
                if (!string.IsNullOrEmpty(exam.SecondTeacher))
                {
                    teacherDisplay = $"{GetLastName(exam.FirstTeacher)}/{GetLastName(exam.SecondTeacher)}";
                }

                // Используем только нужные поля
                string[] data = {
                    exam.ExamTime ?? "",
                    exam.Group ?? "",
                    exam.Subject ?? "",
                    teacherDisplay, // Используем форматированную строку преподавателя (только фамилии)
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

            // Извлекаем только фамилию (первое слово до пробела)
            string[] nameParts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return nameParts.Length > 0 ? nameParts[0] : fullName;
        }

        private string GetDayOfWeekFromDate(string dateString)
        {
            try
            {
                if (DateTime.TryParse(dateString, out DateTime date))
                {
                    // Возвращаем день недели на русском
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

        // Остальные методы остаются без изменений
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

                // Устанавливаем текст с табуляцией в начале
                paragraph.Range.Text = "\t" + lines[i];

                // Безопасное форматирование шрифта
                SafeSetFont(paragraph.Range.Font, "Times New Roman", 12, 1);

                // Форматирование абзаца
                paragraph.Format.SpaceAfter = 0;
                paragraph.Format.SpaceBefore = (i == 0) ? 12 : 0;
                paragraph.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                // Устанавливаем табуляцию
                paragraph.Format.TabStops.Add(_wordApp.CentimetersToPoints(9.5f));

                // Завершаем абзац
                paragraph.Range.InsertParagraphAfter();
            }
        }

        private void CreateDocumentTitle(string titleText)
        {
            Word.Paragraph titleParagraph = _doc.Content.Paragraphs.Add();

            // Устанавливаем текст названия
            titleParagraph.Range.Text = titleText;

            // Безопасное форматирование шрифта
            SafeSetFont(titleParagraph.Range.Font, "Times New Roman", 18, 1);

            // Форматирование абзаца
            titleParagraph.Format.SpaceAfter = 0;
            titleParagraph.Format.SpaceBefore = 18;
            titleParagraph.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            titleParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // Завершаем абзац
            titleParagraph.Range.InsertParagraphAfter();
        }

        private void CreateSimpleTable()
        {
            try
            {
                // Используем подход из вашего старого работающего кода
                Word.Table table = _doc.Tables.Add(
                    _doc.Range(_doc.Content.End - 1),
                    7, // 1 заголовок + 2 дня + 4 строки данных
                    5,
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior,
                    Word.WdAutoFitBehavior.wdAutoFitWindow
                );

                // Убираем границы таблицы
                table.Borders.Enable = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

                // Устанавливаем повторение заголовков на каждой странице
                table.Rows[1].HeadingFormat = -1;

                // Заголовки
                string[] headers = { "Время", "Группа", "Дисциплина", "Преподаватель", "Тип" };
                for (int i = 0; i < headers.Length; i++)
                {
                    table.Cell(1, i + 1).Range.Text = headers[i];
                    table.Cell(1, i + 1).Range.Font.Bold = 1;
                }

                // Заполняем данные напрямую, как в старом коде
                // Понедельник (объединенная строка)
                table.Cell(2, 1).Range.Text = "Понедельник";
                table.Cell(2, 1).Merge(table.Cell(2, 5));
                table.Cell(2, 1).Range.Font.Bold = 1;
                table.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // Белый фон
                table.Cell(2, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;

                // Данные понедельника
                table.Cell(3, 1).Range.Text = "9:00-10:30";
                table.Cell(3, 2).Range.Text = "Группа 101";
                table.Cell(3, 3).Range.Text = "Математика";
                table.Cell(3, 4).Range.Text = "Иванов"; // Только фамилия
                table.Cell(3, 5).Range.Text = "Лекция";

                table.Cell(4, 1).Range.Text = "10:40-12:10";
                table.Cell(4, 2).Range.Text = "Группа 102";
                table.Cell(4, 3).Range.Text = "Физика";
                table.Cell(4, 4).Range.Text = "Петров/Сидоров"; // Два преподавателя
                table.Cell(4, 5).Range.Text = "Практика";

                // Вторник (объединенная строка)
                table.Cell(5, 1).Range.Text = "Вторник";
                table.Cell(5, 1).Merge(table.Cell(5, 5));
                table.Cell(5, 1).Range.Font.Bold = 1;
                table.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                // Белый фон
                table.Cell(5, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;

                // Данные вторника
                table.Cell(6, 1).Range.Text = "13:00-14:30";
                table.Cell(6, 2).Range.Text = "Группа 103";
                table.Cell(6, 3).Range.Text = "Информатика";
                table.Cell(6, 4).Range.Text = "Кузнецов"; // Только фамилия
                table.Cell(6, 5).Range.Text = "Лабораторная";

                // Применяем базовое форматирование ко всей таблице
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        cell.Range.Font.Name = "Times New Roman";
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        // Для заголовков и дней - размер 12-14, для данных - 11
                        if (row.Index == 1) // заголовки
                        {
                            cell.Range.Font.Size = 12;
                            cell.Range.Font.Bold = 1;
                        }
                        else if (row.Index == 2 || row.Index == 5) // дни
                        {
                            cell.Range.Font.Size = 14;
                            cell.Range.Font.Bold = 1;
                        }
                        else // данные
                        {
                            cell.Range.Font.Size = 11;
                            cell.Range.Font.Bold = 0;
                        }
                    }
                }

                _doc.Range(_doc.Content.End - 1).InsertParagraphAfter();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка в создании таблицы: {ex.Message}");
            }
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
                // Если форматирование не удалось, хотя бы текст останется
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
                // Если Times New Roman недоступен, используем Arial
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

        private string GetDocumentsFolderPath()
        {
            string documentsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents");
            if (!Directory.Exists(documentsPath))
            {
                Directory.CreateDirectory(documentsPath);
            }
            return documentsPath;
        }

        private string ShowSaveFileDialog(string initialDirectory)
        {
            var saveFileDialog = new SaveFileDialog
            {
                InitialDirectory = initialDirectory,
                FileName = $"Экзамены_{DateTime.Now:dd.MM.yyyy}",
                Filter = "Word Документы (*.docx)|*.docx|Все файлы (*.*)|*.*",
                DefaultExt = ".docx",
                AddExtension = true,
                Title = "Сохранить документ Word"
            };

            return saveFileDialog.ShowDialog() == true ? saveFileDialog.FileName : null;
        }
    }
}