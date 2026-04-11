using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace ExamScheduleApp.Utilities
{
    public class SimpleDatabaseHelper
    {
        #region СЕТЕВАЯ БД (справочники)
        public string GetConnectionString()
        {
            string dbPath = ConfigurationManager.AppSettings["DatabasePath"];
            return $"Data Source={dbPath};Version=3;Journal Mode=Delete;Pooling=False;BusyTimeout=30000;Synchronous=Normal;Cache=Shared;";
        }

        public void CheckDatabaseStructure()
        {
            try
            {
                Logger.Info("Начата проверка структуры сетевой БД");

                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();

                    using (var cmd = new SQLiteCommand(connection))
                    {
                        cmd.CommandText = "PRAGMA journal_mode=DELETE;";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "PRAGMA busy_timeout=30000;";
                        cmd.ExecuteNonQuery();
                    }

                    CheckTableStructure(connection, "Teachers");
                    CheckTableStructure(connection, "Groups");
                    CheckTableStructure(connection, "Disciplines");
                }

                Logger.Info("Проверка структуры сетевой БД успешно завершена");

            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка при проверке структуры сетевой БД", ex);
                MessageBox.Show($"Ошибка проверки структуры сетевой БД: {ex.Message}");
            }
        }

        private void CheckTableStructure(SQLiteConnection connection, string tableName)
        {
            try
            {
                Logger.Info($"Начата проверка структуры таблиц сетевой БД");

                string query = $"PRAGMA table_info({tableName});";
                using (var command = new SQLiteCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read()) { }
                }

                Logger.Info("Проверка структуры таблиц сетевой БД успешно завершена");

            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка при проверке таблицы {tableName}", ex);
                MessageBox.Show($"Ошибка проверки таблицы {tableName}: {ex.Message}");
            }
        }
        #endregion

        #region ЛОКАЛЬНАЯ БД (расписание)
        public string GetLocalDatabasePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dataFolder = Path.Combine(baseDir, "Data");
            Directory.CreateDirectory(dataFolder);

            return Path.Combine(dataFolder, "LocalExams.db");
        }

        public string GetLocalConnectionString()
        {
            string dbPath = GetLocalDatabasePath();
            return $"Data Source={dbPath};Version=3;Journal Mode=Delete;Pooling=False;BusyTimeout=30000;Synchronous=Normal;Cache=Shared;";
        }

        public void CheckLocalDatabaseStructure()
        {
            try
            {
                // ДЛЯ ОТЛАДКИ
                //string dbPath = GetLocalDatabasePath();
                //MessageBox.Show($"Локальная БД используется по пути:\n{dbPath}", "Информация", MessageBoxButton.OK, MessageBoxImage.Information); // для отладки
                //===================================================

                Logger.Info("Начата проверка/создание локальной БД");

                using (var connection = new SQLiteConnection(GetLocalConnectionString()))
                {
                    connection.Open();

                    string createTable = @"
                        CREATE TABLE IF NOT EXISTS Exams (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            Teacher1Id INTEGER NOT NULL,
                            Teacher2Id INTEGER,
                            SubjectId INTEGER NOT NULL,
                            GroupId INTEGER NOT NULL,
                            Classroom TEXT,
                            Department TEXT,
                            ExamDate TEXT,
                            ExamTime TEXT,
                            ExamType TEXT
                        );";

                    using (var cmd = new SQLiteCommand(createTable, connection))
                        cmd.ExecuteNonQuery();

                    var columns = new List<string>();
                    using (var cmd = new SQLiteCommand("PRAGMA table_info(Exams);", connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read()) columns.Add(reader.GetString(1));
                    }

                    if (!columns.Contains("ExamDate"))
                        new SQLiteCommand("ALTER TABLE Exams ADD COLUMN ExamDate TEXT;", connection).ExecuteNonQuery();
                    if (!columns.Contains("ExamTime"))
                        new SQLiteCommand("ALTER TABLE Exams ADD COLUMN ExamTime TEXT;", connection).ExecuteNonQuery();
                    if (!columns.Contains("ExamType"))
                        new SQLiteCommand("ALTER TABLE Exams ADD COLUMN ExamType TEXT;", connection).ExecuteNonQuery();
                }

                Logger.Info("Локальная база данных успешно проверена/создана");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка работы с локальной БД:\n{ex.Message}\n\nПуть: {GetLocalDatabasePath()}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Logger.Error("Ошибка CheckLocalDatabaseStructure", ex);
            }
        }
        #endregion

        #region РАБОТА С РАСПИСАНИЕМ (ЛОКАЛЬНАЯ БД)
        public List<ExamSchedule> GetExamSchedule()
        {
            var exams = new List<ExamSchedule>();
            try
            {
                Logger.Info("Начата загрузка экзаменов из локальной БД");

                using (var connection = new SQLiteConnection(GetLocalConnectionString()))
                {
                    connection.Open();
                    string query = @"SELECT Id, Teacher1Id, Teacher2Id, SubjectId, GroupId, 
                                           Classroom, Department, ExamDate, ExamTime, ExamType 
                                    FROM Exams";

                    using (var command = new SQLiteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var exam = new ExamSchedule
                            {
                                Id = SafeGetInt32(reader, "Id"),
                                Teacher1Id = SafeGetInt32(reader, "Teacher1Id"),
                                Teacher2Id = SafeGetNullableInt32(reader, "Teacher2Id"),
                                SubjectId = SafeGetInt32(reader, "SubjectId"),
                                GroupId = SafeGetInt32(reader, "GroupId"),
                                Classroom = SafeGetString(reader, "Classroom"),
                                DepartmentName = SafeGetString(reader, "Department"),
                                ExamDate = SafeGetString(reader, "ExamDate"),
                                ExamTime = SafeGetString(reader, "ExamTime"),
                                ExamType = SafeGetString(reader, "ExamType")
                            };

                            exam.Teacher1Name = GetTeacherName(exam.Teacher1Id);
                            exam.Teacher2Name = exam.Teacher2Id.HasValue ? GetTeacherName(exam.Teacher2Id.Value) : null;
                            exam.SubjectName = GetSubjectName(exam.SubjectId);
                            exam.GroupName = GetGroupName(exam.GroupId);

                            exams.Add(exam);
                        }
                    }
                }

                Logger.Info($"Загружено {exams.Count} экзаменов из локальной БД");

            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка загрузки экзаменов", ex);
                MessageBox.Show($"Ошибка загрузки экзаменов: {ex.Message}");
                return new List<ExamSchedule>();
            }
            return exams;
        }

        public void AddExam(ExamSchedule exam, bool showInfo = false)
        {
            try
            {
                using (var connection = new SQLiteConnection(GetLocalConnectionString()))
                {
                    connection.Open();
                    string query = @"
                        INSERT INTO Exams (Teacher1Id, Teacher2Id, SubjectId, GroupId, Classroom, 
                                           Department, ExamDate, ExamTime, ExamType) 
                        VALUES (@Teacher1Id, @Teacher2Id, @SubjectId, @GroupId, @Classroom, 
                                @Department, @ExamDate, @ExamTime, @ExamType)";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Teacher1Id", exam.Teacher1Id);
                        command.Parameters.AddWithValue("@Teacher2Id", exam.Teacher2Id.HasValue ? (object)exam.Teacher2Id.Value : DBNull.Value);
                        command.Parameters.AddWithValue("@SubjectId", exam.SubjectId);
                        command.Parameters.AddWithValue("@GroupId", exam.GroupId);
                        command.Parameters.AddWithValue("@Classroom", exam.Classroom ?? "");
                        command.Parameters.AddWithValue("@Department", exam.DepartmentName ?? "");
                        command.Parameters.AddWithValue("@ExamDate", exam.ExamDate ?? "");
                        command.Parameters.AddWithValue("@ExamTime", exam.ExamTime ?? "");
                        command.Parameters.AddWithValue("@ExamType", exam.ExamType ?? "");

                        command.ExecuteNonQuery();
                    }
                }

                Logger.Info($"Добавлен экзамен: {exam.SubjectName} | {exam.GroupName} | {exam.ExamDate}");
                if (showInfo) MessageBox.Show("Экзамен успешно добавлен!");

            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка AddExam", ex);
                MessageBox.Show($"Ошибка добавления экзамена: {ex.Message}");
            }
        }

        public void DeleteExam(int examId, bool showInfo = false)
        {
            try
            {
                Logger.Warning($"Удаление экзамена ID: {examId}");

                using (var connection = new SQLiteConnection(GetLocalConnectionString()))
                {
                    connection.Open();
                    string query = "DELETE FROM Exams WHERE Id = @Id";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", examId);
                        int rowsDeleted = command.ExecuteNonQuery();

                        if (showInfo)
                        {
                            if (rowsDeleted > 0)
                            {
                                Logger.Info($"Экзамен ID {examId} успешно удалён");
                                MessageBox.Show("Экзамен успешно удален из базы данных");
                            }

                            else
                            {
                                Logger.Warning($"Экзамен ID {examId} не найден");
                                MessageBox.Show("Экзамен не найден");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка удаления экзамена ID {examId}", ex);
                MessageBox.Show($"Ошибка удаления экзамена: {ex.Message}");
            }
        }

        public void ClearAllExams()
        {
            try
            {
                Logger.Warning("Запрошена полная очистка всех экзаменов в локальной БД");

                using (var connection = new SQLiteConnection(GetLocalConnectionString()))
                {
                    connection.Open();
                    string query = "DELETE FROM Exams";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.ExecuteNonQuery();
                        Logger.Info("Все экзамены из локальной БД были успешно удалены");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка в очистке всего списка экзаменов", ex);
                MessageBox.Show($"Ошибка очистки экзаменов: {ex.Message}");
            }
        }
        #endregion

        #region ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
        private string SafeGetString(SQLiteDataReader reader, string columnName)
        {
            try
            {
                int columnIndex = reader.GetOrdinal(columnName);
                if (!reader.IsDBNull(columnIndex))
                    return reader.GetString(columnIndex);
                return string.Empty;
            }
            catch { return string.Empty; }
        }

        private int SafeGetInt32(SQLiteDataReader reader, string columnName)
        {
            try
            {
                int columnIndex = reader.GetOrdinal(columnName);
                if (!reader.IsDBNull(columnIndex))
                    return reader.GetInt32(columnIndex);
                return 0;
            }
            catch { return 0; }
        }

        private int? SafeGetNullableInt32(SQLiteDataReader reader, string columnName)
        {
            try
            {
                int columnIndex = reader.GetOrdinal(columnName);
                if (!reader.IsDBNull(columnIndex))
                    return reader.GetInt32(columnIndex);
                return null;
            }
            catch { return null; }
        }

        private string GetTeacherName(int teacherId)
        {
            try
            {
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT name FROM Teachers WHERE id = @id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@id", teacherId);
                        return command.ExecuteScalar()?.ToString() ?? "";
                    }
                }
            }
            catch { return ""; }
        }

        private string GetSubjectName(int subjectId)
        {
            try
            {
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT shortName9 FROM Disciplines WHERE id = @id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@id", subjectId);
                        return command.ExecuteScalar()?.ToString() ?? "";
                    }
                }
            }
            catch { return ""; }
        }

        private string GetGroupName(int groupId)
        {
            try
            {
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT name FROM Groups WHERE id = @id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@id", groupId);
                        return command.ExecuteScalar()?.ToString() ?? "";
                    }
                }
            }
            catch { return ""; }
        }
        #endregion

        #region МЕТОДЫ СПРАВОЧНИКОВ
        public List<Teacher> GetTeachers()
        {
            var teachers = new List<Teacher>();
            try
            {
                Logger.Info("Загрузка списка преподавателей...");

                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT id, name, classroom, academicBuilding FROM Teachers ORDER BY name";

                    using (var command = new SQLiteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            teachers.Add(new Teacher
                            {
                                Id = SafeGetInt32(reader, "id"),
                                Name = SafeGetString(reader, "name"),
                                Classroom = SafeGetString(reader, "classroom"),
                                AcademicBuilding = SafeGetInt32(reader, "academicBuilding")
                            });
                        }
                    }
                }

                Logger.Info($"Успешно загружено {teachers.Count} преподавателей");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка загрузки преподавателей", ex);
                MessageBox.Show($"Ошибка загрузки преподавателей: {ex.Message}");
            }
            return teachers;
        }

        public List<Subject> GetSubjects()
        {
            var subjects = new List<Subject>();
            try
            {
                Logger.Info("Загрузка списка дисциплин...");

                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT id, fullname, shortName12, shortName9, shortName5 FROM Disciplines ORDER BY fullname";

                    using (var command = new SQLiteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            subjects.Add(new Subject
                            {
                                Id = SafeGetInt32(reader, "id"),
                                FullName = SafeGetString(reader, "fullname"),
                                ShortName12 = SafeGetString(reader, "shortName12"),
                                ShortName9 = SafeGetString(reader, "shortName9"),
                                ShortName5 = SafeGetString(reader, "shortName5")
                            });
                        }
                    }
                }

                Logger.Info($"Успешно загружено {subjects.Count} дисциплин");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка загрузки дисциплин", ex);
                MessageBox.Show($"Ошибка загрузки дисциплин: {ex.Message}");
            }
            return subjects;
        }

        public List<Group> GetGroups()
        {
            var groups = new List<Group>();
            try
            {
                Logger.Info("Загрузка списка групп...");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "SELECT id, name, department FROM Groups ORDER BY name";

                    using (var command = new SQLiteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            groups.Add(new Group
                            {
                                Id = SafeGetInt32(reader, "id"),
                                Name = SafeGetString(reader, "name"),
                                Department = SafeGetString(reader, "department")
                            });
                        }
                    }
                }
                Logger.Info($"Успешно загружено {groups.Count} групп");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка загрузки групп", ex);
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}");
            }
            return groups;
        }

        public void AddTeacher(Teacher teacher)
        {
            try
            {
                Logger.Info($"Добавление преподавателя: {teacher.Name}");
                string query = "INSERT INTO Teachers (name, classroom, academicBuilding) VALUES (@Name, @Classroom, @AcademicBuilding)";
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Name", teacher.Name);
                        command.Parameters.AddWithValue("@Classroom", teacher.Classroom);
                        command.Parameters.AddWithValue("@AcademicBuilding", teacher.AcademicBuilding);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Преподаватель {teacher.Name} успешно добавлен");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка добавления преподавателя {teacher?.Name}", ex);
                MessageBox.Show($"Ошибка добавления преподавателя: {ex.Message}");
            }
        }

        public void AddSubject(Subject subject)
        {
            try
            {
                Logger.Info($"Добавление дисциплины: {subject.ShortName9}");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "INSERT INTO Disciplines (fullname, shortName12, shortName9, shortName5) VALUES (@FullName, @ShortName12, @ShortName9, @ShortName5)";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        string shortName9 = subject.ShortName9 ?? "";
                        command.Parameters.AddWithValue("@FullName", shortName9);
                        command.Parameters.AddWithValue("@ShortName12", shortName9.Length > 12 ? shortName9.Substring(0, 12) : shortName9);
                        command.Parameters.AddWithValue("@ShortName9", shortName9);
                        command.Parameters.AddWithValue("@ShortName5", shortName9.Length > 5 ? shortName9.Substring(0, 5) : shortName9);

                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Дисциплина '{subject.ShortName9}' успешно добавлена");
                MessageBox.Show($"Дисциплина '{subject.ShortName9}' успешно добавлена!");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка добавления дисциплины {subject?.ShortName9}", ex);
                MessageBox.Show($"Ошибка добавления дисциплины: {ex.Message}");
            }
        }

        public void AddGroup(Group group)
        {
            try
            {
                Logger.Info($"Добавление группы: {group.Name}");
                string query = "INSERT INTO Groups (name, department) VALUES (@Name, @Department)";
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Name", group.Name);
                        command.Parameters.AddWithValue("@Department", group.Department ?? "");
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Группа {group.Name} успешно добавлена");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка добавления группы {group?.Name}", ex);
                MessageBox.Show($"Ошибка добавления группы: {ex.Message}");
            }
        }

        public bool TeacherExists(string name)
        {
            string query = "SELECT COUNT(*) FROM Teachers WHERE name = @Name";
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Name", name);
                    var count = Convert.ToInt32(command.ExecuteScalar());
                    return count > 0;
                }
            }
        }

        public bool SubjectExists(string shortName9)
        {
            string query = "SELECT COUNT(*) FROM Disciplines WHERE shortName9 = @ShortName9";
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ShortName9", shortName9);
                    var count = Convert.ToInt32(command.ExecuteScalar());
                    return count > 0;
                }
            }
        }

        public bool GroupExists(string name)
        {
            string query = "SELECT COUNT(*) FROM Groups WHERE name = @Name";
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Name", name);
                    var count = Convert.ToInt32(command.ExecuteScalar());
                    return count > 0;
                }
            }
        }

        public void UpdateTeacher(Teacher teacher)
        {
            try
            {
                Logger.Info($"Обновление преподавателя: {teacher.Name} (ID: {teacher.Id})");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "UPDATE Teachers SET name = @Name, classroom = @Classroom, academicBuilding = @AcademicBuilding WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Name", teacher.Name);
                        command.Parameters.AddWithValue("@Classroom", teacher.Classroom);
                        command.Parameters.AddWithValue("@AcademicBuilding", teacher.AcademicBuilding);
                        command.Parameters.AddWithValue("@Id", teacher.Id);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Преподаватель {teacher.Name} успешно обновлён");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка обновления преподавателя {teacher?.Name}", ex);
                MessageBox.Show($"Ошибка обновления преподавателя: {ex.Message}");
            }
        }

        public void DeleteTeacher(int teacherId)
        {
            try
            {
                Logger.Warning($"Попытка удаления преподавателя ID: {teacherId}");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "DELETE FROM Teachers WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", teacherId);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Преподаватель ID {teacherId} успешно удалён");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка удаления преподавателя ID {teacherId}", ex);
                MessageBox.Show($"Ошибка удаления преподавателя: {ex.Message}");
            }
        }

        public void UpdateSubject(Subject subject)
        {
            try
            {
                Logger.Info($"Обновление дисциплины: {subject.ShortName9} (ID: {subject.Id})");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "UPDATE Disciplines SET fullname = @FullName, shortName12 = @ShortName12, shortName9 = @ShortName9, shortName5 = @ShortName5 WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@FullName", subject.FullName);
                        command.Parameters.AddWithValue("@ShortName12", subject.ShortName12);
                        command.Parameters.AddWithValue("@ShortName9", subject.ShortName9);
                        command.Parameters.AddWithValue("@ShortName5", subject.ShortName5);
                        command.Parameters.AddWithValue("@Id", subject.Id);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Дисциплина {subject.ShortName9} успешно обновлена");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка обновления дисциплины {subject?.ShortName9}", ex);
                MessageBox.Show($"Ошибка обновления дисциплины: {ex.Message}");
            }
        }

        public void DeleteSubject(int subjectId)
        {
            try
            {
                Logger.Warning($"Попытка удаления дисциплины ID: {subjectId}");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "DELETE FROM Disciplines WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", subjectId);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Дисциплина ID {subjectId} успешно удалена");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка удаления дисциплины ID {subjectId}", ex);
                MessageBox.Show($"Ошибка удаления дисциплины: {ex.Message}");
            }
        }

        public void UpdateGroup(Group group)
        {
            try
            {
                Logger.Info($"Обновление группы: {group.Name} (ID: {group.Id})");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "UPDATE Groups SET name = @Name, department = @Department WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Name", group.Name);
                        command.Parameters.AddWithValue("@Department", group.Department);
                        command.Parameters.AddWithValue("@Id", group.Id);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Группа {group.Name} успешно обновлена");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка обновления группы {group?.Name}", ex);
                MessageBox.Show($"Ошибка обновления группы: {ex.Message}");
            }
        }

        public void DeleteGroup(int groupId)
        {
            try
            {
                Logger.Warning($"Попытка удаления группы ID: {groupId}");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = "DELETE FROM Groups WHERE id = @Id";
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Id", groupId);
                        command.ExecuteNonQuery();
                    }
                }
                Logger.Info($"Группа ID {groupId} успешно удалена");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка удаления группы ID {groupId}", ex);
                MessageBox.Show($"Ошибка удаления группы: {ex.Message}");
            }
        }

        public void CleanProblematicData()
        {
            try
            {
                Logger.Info("Начата очистка проблемных данных в справочниках");
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string[] cleanupQueries = {
                        "DELETE FROM Teachers WHERE name IS NULL OR classroom IS NULL",
                        "DELETE FROM Groups WHERE name IS NULL",
                        "DELETE FROM Disciplines WHERE shortName9 IS NULL"
                    };

                    foreach (string query in cleanupQueries)
                    {
                        using (var command = new SQLiteCommand(query, connection))
                        {
                            int affected = command.ExecuteNonQuery();
                            if (affected > 0)
                                Logger.Info($"Удалено {affected} проблемных записей в справочниках");
                        }
                    }
                }
                Logger.Info("Очистка проблемных данных завершена");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка очистки проблемных данных", ex);
                MessageBox.Show($"Ошибка очистки данных: {ex.Message}");
            }
        }
        #endregion
    }
}