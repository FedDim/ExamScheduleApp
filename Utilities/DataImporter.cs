using ClosedXML.Excel;
using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExamScheduleApp.Utilities
{
    public class ImportResult
    {
        public int Added { get; set; }
        public int Skipped { get; set; }
        public int ErrorsCount { get; set; }
        public List<string> ErrorMessages { get; set; } = new List<string>();

        public string GetSummary()
        {
            return $"Импорт завершён:\n\n" +
                   $"Добавлено новых записей: {Added}\n" +
                   $"Пропущено (уже существуют): {Skipped}\n" +
                   $"Ошибок: {ErrorsCount}";
        }
    }

    public class DataImporter
    {
        private readonly SimpleDatabaseHelper _db;

        public DataImporter(SimpleDatabaseHelper dbHelper)
        {
            _db = dbHelper;
        }

        public ImportResult Import(DataType dataType, string excelFilePath)
        {
            var result = new ImportResult();

            if (!File.Exists(excelFilePath))
            {
                result.ErrorMessages.Add("Файл не найден.");
                result.ErrorsCount = 1;
                return result;
            }

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    int rowCount = worksheet.LastRowUsed()?.RowNumber() ?? 0;

                    for (int rowNum = 2; rowNum <= rowCount; rowNum++)
                    {
                        var row = worksheet.Row(rowNum);
                        try
                        {
                            switch (dataType)
                            {
                                case DataType.TEACHER:
                                    ImportTeacher(row, result);
                                    break;
                                case DataType.GROUP:
                                    ImportGroup(row, result);
                                    break;
                                case DataType.SUBJECT:
                                    ImportSubject(row, result);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            result.ErrorsCount++;
                            result.ErrorMessages.Add($"Строка {rowNum}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка импорта {dataType}", ex);
                result.ErrorMessages.Add($"Ошибка чтения файла: {ex.Message}");
                result.ErrorsCount++;
            }

            return result;
        }

        private void ImportTeacher(IXLRow row, ImportResult result)
        {
            string name = row.Cell(1).GetString().Trim();
            if (string.IsNullOrEmpty(name)) return;

            if (_db.TeacherExists(name))
            {
                result.Skipped++;
                return;
            }

            var teacher = new Teacher
            {
                Name = name,
                Classroom = row.Cell(2).GetString().Trim(),
                AcademicBuilding = int.TryParse(row.Cell(3).GetString(), out int b) ? b : 0
            };

            _db.AddTeacher(teacher);
            result.Added++;
        }

        private void ImportGroup(IXLRow row, ImportResult result)
        {
            string name = row.Cell(1).GetString().Trim();
            if (string.IsNullOrEmpty(name)) return;

            if (_db.GroupExists(name))
            {
                result.Skipped++;
                return;
            }

            var group = new Group
            {
                Name = name,
                Department = row.Cell(2).GetString().Trim()
            };

            _db.AddGroup(group);
            result.Added++;
        }

        private void ImportSubject(IXLRow row, ImportResult result)
        {
            string shortName9 = row.Cell(1).GetString().Trim();
            if (string.IsNullOrEmpty(shortName9)) return;

            if (_db.SubjectExists(shortName9))
            {
                result.Skipped++;
                return;
            }

            var subject = new Subject
            {
                ShortName9 = shortName9,
                FullName = row.Cell(2).GetString().Trim(),
                ShortName12 = row.Cell(3).GetString().Trim(),
                ShortName5 = row.Cell(4).GetString().Trim()
            };

            _db.AddSubject(subject);
            result.Added++;
        }
    }
}