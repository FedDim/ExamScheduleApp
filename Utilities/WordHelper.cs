using ExamScheduleApp.Model;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
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
            _wordApp = new Word.Application();
            _doc = _wordApp.Documents.Add();
            _exams = exams;
        }

        public void CreateDocument()
        {
            try
            {
                // Настройка полей документа (в пунктах) ~3 см
                _doc.PageSetup.LeftMargin = 85;
                _doc.PageSetup.RightMargin = 85;
                _doc.PageSetup.TopMargin = 85;
                _doc.PageSetup.BottomMargin = 85;

                // Создание шапки документа
                CreateHeader(_doc);

                // Добавление таблицы с данными
                CreateExamTable(_doc);

                string documentsPath = GetDocumentsFolderPath();

                string filePath = ShowSaveFileDialog(documentsPath);

                if (!string.IsNullOrEmpty(filePath))
                {
                    _doc.SaveAs2(filePath);
                    MessageBox.Show($"Файл успешно сохранён по пути : {filePath}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (_doc != null)
                {
                    _doc.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_doc);
                }

                if (_wordApp != null)
                {
                    _wordApp.Quit(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wordApp);
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
            int rowCount = _exams.Count;
            int columnCount = 8; // Преподаватель 1, Преподаватель 2, Дата, Дисциплина, Группа, Время, Аудитория, Тип

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
            wordTable.Range.Font.Bold = 0;
            wordTable.Range.Font.Italic = 0;
            wordTable.Range.Font.Color = Word.WdColor.wdColorBlack;

            wordTable.Borders.Enable = 0;
            wordTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
            wordTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;

            // Заполняем заголовки
            string[] headers = { "Преподаватель 1", "Преподаватель 2", "Дата", "Дисциплина", "Группа", "Время", "Аудитория", "Тип" };
            for (int col = 0; col < columnCount; col++)
            {
                wordTable.Cell(1, col + 1).Range.Text = headers[col];
                // Делаем заголовки жирными
                wordTable.Cell(1, col + 1).Range.Font.Bold = 1;
            }

            // Заполняем данные из коллекции _exams
            for (int row = 0; row < rowCount; row++)
            {
                var exam = _exams[row];

                wordTable.Cell(row + 2, 1).Range.Text = exam.FirstTeacher ?? "";
                wordTable.Cell(row + 2, 2).Range.Text = exam.SecondTeacher ?? "";
                wordTable.Cell(row + 2, 3).Range.Text = exam.ExamDate ?? "";
                wordTable.Cell(row + 2, 4).Range.Text = exam.Subject ?? "";
                wordTable.Cell(row + 2, 5).Range.Text = exam.Group ?? "";
                wordTable.Cell(row + 2, 6).Range.Text = exam.ExamTime ?? "";
                wordTable.Cell(row + 2, 7).Range.Text = exam.Classroom ?? "";
                wordTable.Cell(row + 2, 8).Range.Text = exam.ExamType ?? "";
            }

            // Форматирование таблицы
            wordTable.Range.Font.Size = 10;

            // Добавляем отступ после таблицы
            doc.Range(doc.Content.End - 1).InsertParagraphAfter();
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
    }
}
