using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace ExamScheduleApp.Utilities
{
    public class SimpleDatabaseHelper
    {
        // ====================== ПОДКЛЮЧЕНИЯ ======================
        private readonly string _serverConnectionString;   // SQL Server — справочники
        private readonly string _localConnectionString;    // SQLite — расписание

        public SimpleDatabaseHelper()
        {
            try
            {
                _serverConnectionString = ConfigurationManager.ConnectionStrings["ExamScheduleServer"].ConnectionString;

                string localPath = GetLocalDatabasePath();
                _localConnectionString = $"Data Source={localPath};Version=3;Journal Mode=Delete;Pooling=False;BusyTimeout=30000;";

                Logger.Info("SimpleDatabaseHelper: Строки подключения загружены.");
            }
            catch (Exception ex)
            {
                Logger.Error("Критическая ошибка при чтении App.config", ex);
                throw;
            }
        }

        private string GetLocalDatabasePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dataFolder = Path.Combine(baseDir, "Data");
            Directory.CreateDirectory(dataFolder);
            return Path.Combine(dataFolder, "LocalExams.db");
        }

        // ====================== ИНИЦИАЛИЗАЦИЯ ======================
        public bool CheckServerConnection()
        {
            using (var conn = new SqlConnection(_serverConnectionString))
            {
                try
                {
                    Logger.Info($"Попытка подключения к серверу: {conn.DataSource}");
                    conn.Open();
                    Logger.Info("Сетевая база данных доступна.");
                    return true;
                }
                catch (SqlException ex)
                {
                    // Детальный разбор ошибок SQL
                    string errorMsg = "Ошибка подключения к SQL Server:\n";
                    switch (ex.Number)
                    {
                        case -1:
                        case 2:
                        case 53:
                            errorMsg += "Сервер не найден или недоступен. Проверьте IP адрес и Firewall.";
                            break;
                        case 4060:
                            errorMsg += "База данных ExamScheduleDB не найдена на сервере.";
                            break;
                        case 18456:
                            errorMsg += "Ошибка входа. Проверьте логин (ExamAppUser) и пароль.";
                            break;
                        default:
                            errorMsg += ex.Message;
                            break;
                    }
                    Logger.Error(errorMsg, ex);
                    MessageBox.Show(errorMsg, "Ошибка сети", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }
        }

        public void CheckLocalDatabaseStructure()
        {
            try
            {
                using (var conn = new SQLiteConnection(_localConnectionString))
                {
                    conn.Open();
                    string sql = @"
                CREATE TABLE IF NOT EXISTS LocalExams (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Teacher1Id INTEGER,
                    Teacher2Id INTEGER,
                    SubjectId INTEGER,
                    GroupId INTEGER,
                    ExamDate TEXT,
                    ExamTime TEXT,
                    Classroom TEXT,
                    ExamType TEXT
                )";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                Logger.Info("Структура локальной БД проверена (LocalExams)");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка CheckLocalDatabaseStructure", ex);
            }
        }

        // ====================== СПРАВОЧНИКИ (SQL Server) ======================
        public List<Teacher> GetTeachers()
        {
            var list = new List<Teacher>();
            using (var conn = new SqlConnection(_serverConnectionString))
            {
                try
                {
                    conn.Open();
                    string sql = "SELECT Id, Name, Classroom, AcademicBuilding FROM Teachers ORDER BY Name";
                    using (var cmd = new SqlCommand(sql, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new Teacher
                            {
                                Id = reader.GetInt32(0),
                                Name = reader.IsDBNull(1) ? "" : reader.GetString(1),
                                Classroom = reader.IsDBNull(2) ? "" : reader.GetString(2),
                                AcademicBuilding = reader.IsDBNull(3) ? 0 : reader.GetInt32(3)
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Ошибка загрузки преподавателей", ex);
                }
            }
            return list;
        }

        public List<Subject> GetSubjects()
        {
            var list = new List<Subject>();
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT Id, FullName, ShortName12, ShortName9, ShortName5 FROM Disciplines ORDER BY ShortName9", conn))
                {
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            list.Add(new Subject
                            {
                                Id = r.GetInt32(0),
                                FullName = r.IsDBNull(1) ? "" : r.GetString(1),
                                ShortName12 = r.IsDBNull(2) ? "" : r.GetString(2),
                                ShortName9 = r.GetString(3),
                                ShortName5 = r.IsDBNull(4) ? "" : r.GetString(4)
                            });
                        }
                    }
                }
                Logger.Info($"Загружено {list.Count} дисциплин");
            }
            catch (Exception ex) { Logger.Error("GetSubjects", ex); MessageBox.Show($"Ошибка дисциплин: {ex.Message}"); }
            return list;
        }

        public List<Group> GetGroups()
        {
            var list = new List<Group>();
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT Id, Name, Department FROM Groups ORDER BY Name", conn))
                {
                    conn.Open();
                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            list.Add(new Group
                            {
                                Id = r.GetInt32(0),
                                Name = r.GetString(1),
                                Department = r.IsDBNull(2) ? "" : r.GetString(2)
                            });
                        }
                    }
                }
                Logger.Info($"Загружено {list.Count} групп");
            }
            catch (Exception ex) { Logger.Error("GetGroups", ex); MessageBox.Show($"Ошибка групп: {ex.Message}"); }
            return list;
        }

        // ====================== EXISTS МЕТОДЫ ======================
        public bool TeacherExists(string name)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM Teachers WHERE Name = @name", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@name", name);
                    int count = (int)cmd.ExecuteScalar();
                    return count > 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"TeacherExists ({name})", ex);
                return false;
            }
        }

        public bool SubjectExists(string shortName9)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM Disciplines WHERE ShortName9 = @name", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@name", shortName9);
                    int count = (int)cmd.ExecuteScalar();
                    return count > 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"SubjectExists ({shortName9})", ex);
                return false;
            }
        }

        public bool GroupExists(string name)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT COUNT(*) FROM Groups WHERE Name = @name", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@name", name);
                    int count = (int)cmd.ExecuteScalar();
                    return count > 0;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"GroupExists ({name})", ex);
                return false;
            }
        }

        #region ====================== CRUD ДЛЯ СПРАВОЧНИКОВ ======================

        // ==================== TEACHERS ====================
        public void AddTeacher(Teacher teacher)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("INSERT INTO Teachers (Name, Classroom, AcademicBuilding) VALUES (@n, @c, @b)", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@n", teacher.Name);
                    cmd.Parameters.AddWithValue("@c", teacher.Classroom ?? "");
                    cmd.Parameters.AddWithValue("@b", teacher.AcademicBuilding);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Добавлен преподаватель: {teacher.Name}");
            }
            catch (Exception ex) { Logger.Error($"AddTeacher {teacher?.Name}", ex); MessageBox.Show($"Ошибка добавления преподавателя: {ex.Message}"); }
        }

        public void UpdateTeacher(Teacher teacher)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("UPDATE Teachers SET Name=@n, Classroom=@c, AcademicBuilding=@b WHERE Id=@id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@n", teacher.Name);
                    cmd.Parameters.AddWithValue("@c", teacher.Classroom ?? "");
                    cmd.Parameters.AddWithValue("@b", teacher.AcademicBuilding);
                    cmd.Parameters.AddWithValue("@id", teacher.Id);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Обновлён преподаватель ID {teacher.Id}");
            }
            catch (Exception ex) { Logger.Error($"UpdateTeacher {teacher?.Id}", ex); }
        }

        public void DeleteTeacher(int teacherId)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("DELETE FROM Teachers WHERE Id = @id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@id", teacherId);
                    cmd.ExecuteNonQuery();
                }
                Logger.Warning($"Удалён преподаватель ID: {teacherId}");
            }
            catch (Exception ex) { Logger.Error($"DeleteTeacher {teacherId}", ex); }
        }

        /// <summary>
        /// Возвращает аудиторию преподавателя по его ID (из SQL Server)
        /// </summary>
        public string GetTeacherClassroom(int teacherId)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT Classroom FROM Teachers WHERE Id = @id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@id", teacherId);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"GetTeacherClassroom({teacherId})", ex);
                return "";
            }
        }

        /// <summary>
        /// Перегрузка — по имени преподавателя
        /// </summary>
        public string GetTeacherClassroom(string teacherName)
        {
            if (string.IsNullOrWhiteSpace(teacherName)) return "";

            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("SELECT Classroom FROM Teachers WHERE Name = @name", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@name", teacherName);
                    var result = cmd.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"GetTeacherClassroom({teacherName})", ex);
                return "";
            }
        }

        // ==================== SUBJECTS ====================
        public void AddSubject(Subject subject)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("INSERT INTO Disciplines (FullName, ShortName12, ShortName9, ShortName5) VALUES (@f, @s12, @s9, @s5)", conn))
                {
                    conn.Open();
                    string name = subject.ShortName9 ?? "";
                    cmd.Parameters.AddWithValue("@f", name);
                    cmd.Parameters.AddWithValue("@s12", name.Length > 12 ? name.Substring(0, 12) : name);
                    cmd.Parameters.AddWithValue("@s9", name);
                    cmd.Parameters.AddWithValue("@s5", name.Length > 5 ? name.Substring(0, 5) : name);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Добавлена дисциплина: {subject.ShortName9}");
            }
            catch (Exception ex) { Logger.Error($"AddSubject {subject?.ShortName9}", ex); }
        }

        public void UpdateSubject(Subject subject)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("UPDATE Disciplines SET FullName=@f, ShortName12=@s12, ShortName9=@s9, ShortName5=@s5 WHERE Id=@id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@f", subject.FullName);
                    cmd.Parameters.AddWithValue("@s12", subject.ShortName12);
                    cmd.Parameters.AddWithValue("@s9", subject.ShortName9);
                    cmd.Parameters.AddWithValue("@s5", subject.ShortName5);
                    cmd.Parameters.AddWithValue("@id", subject.Id);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Обновлена дисциплина ID {subject.Id}");
            }
            catch (Exception ex) { Logger.Error($"UpdateSubject {subject?.Id}", ex); }
        }

        public void DeleteSubject(int subjectId)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("DELETE FROM Disciplines WHERE Id = @id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@id", subjectId);
                    cmd.ExecuteNonQuery();
                }
                Logger.Warning($"Удалена дисциплина ID: {subjectId}");
            }
            catch (Exception ex) { Logger.Error($"DeleteSubject {subjectId}", ex); }
        }

        // ==================== GROUPS ====================
        public void AddGroup(Group group)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("INSERT INTO Groups (Name, Department) VALUES (@n, @d)", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@n", group.Name);
                    cmd.Parameters.AddWithValue("@d", group.Department ?? "");
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Добавлена группа: {group.Name}");
            }
            catch (Exception ex) { Logger.Error($"AddGroup {group?.Name}", ex); }
        }

        public void UpdateGroup(Group group)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("UPDATE Groups SET Name=@n, Department=@d WHERE Id=@id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@n", group.Name);
                    cmd.Parameters.AddWithValue("@d", group.Department ?? "");
                    cmd.Parameters.AddWithValue("@id", group.Id);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Обновлена группа ID {group.Id}");
            }
            catch (Exception ex) { Logger.Error($"UpdateGroup {group?.Id}", ex); }
        }

        public void DeleteGroup(int groupId)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand("DELETE FROM Groups WHERE Id = @id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@id", groupId);
                    cmd.ExecuteNonQuery();
                }
                Logger.Warning($"Удалена группа ID: {groupId}");
            }
            catch (Exception ex) { Logger.Error($"DeleteGroup {groupId}", ex); }
        }
        #endregion

        // ====================== РАСПИСАНИЕ (SQLite) ======================
        public List<ExamSchedule> GetExamSchedule()
        {
            var exams = new List<ExamSchedule>();
            try
            {
                using (var conn = new SQLiteConnection(_localConnectionString))
                using (var cmd = new SQLiteCommand("SELECT * FROM Exams", conn))
                {
                    conn.Open();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var exam = new ExamSchedule
                            {
                                Id = SafeGetInt32(reader, "Id"),
                                Teacher1Id = SafeGetInt32(reader, "Teacher1Id"),
                                Teacher2Id = SafeGetInt32(reader, "Teacher2Id"),
                                SubjectId = SafeGetInt32(reader, "SubjectId"),
                                GroupId = SafeGetInt32(reader, "GroupId"),
                                Classroom = SafeGetString(reader, "Classroom"),
                                DepartmentName = SafeGetString(reader, "Department")
                            };

                            exam.Teacher1Name = GetTeacherName(exam.Teacher1Id);
                            exam.Teacher2Name = exam.Teacher2Id.HasValue ? GetTeacherName(exam.Teacher2Id.Value) : null;
                            exam.SubjectName = GetSubjectName(exam.SubjectId);
                            exam.GroupName = GetGroupName(exam.GroupId);

                            exams.Add(exam);
                        }
                    }
                }
            }
            catch (Exception ex) { Logger.Error("GetExamSchedule", ex); }
            return exams;
        }

        public void AddExam(ExamSchedule exam, bool showInfo = false)
        {
            try
            {
                using (var connection = new SQLiteConnection(GetConnectionString()))
                {
                    connection.Open();
                    string query = @"
                                    INSERT INTO Exams (Teacher1Id, Teacher2Id, SubjectId, GroupId, 
                                                     Classroom, Department, ExamDate, ExamTime, ExamType) 
                                    VALUES (@Teacher1Id, @Teacher2Id, @SubjectId, @GroupId, 
                                            @Classroom, @Department, @ExamDate, @ExamTime, @ExamType)";
                    using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Teacher1Id", exam.Teacher1Id);
                    command.Parameters.AddWithValue("@Teacher2Id", exam.Teacher2Id);
                    command.Parameters.AddWithValue("@SubjectId", exam.SubjectId);
                    command.Parameters.AddWithValue("@GroupId", exam.GroupId);
                    command.Parameters.AddWithValue("@Classroom", exam.Classroom);
                    command.Parameters.AddWithValue("@Department", exam.DepartmentName);
                    command.Parameters.AddWithValue("@ExamDate", exam.ExamDate);
                    command.Parameters.AddWithValue("@ExamTime", exam.ExamTime);
                    command.Parameters.AddWithValue("@ExamType", exam.ExamType);

                    command.ExecuteNonQuery();
                }
            }
                Logger.Info($"Экзамен сохранен: Дата {exam.ExamDate}, Время {exam.ExamTime}");
                if (showInfo) MessageBox.Show("Данные успешно сохранены!");
            }
            catch (Exception ex)
            {
                Logger.Error("Ошибка в AddExam", ex);
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        public void DeleteExam(int examId, bool showInfo = false)
        {
            try
            {
                using (var conn = new SQLiteConnection(_localConnectionString))
                using (var cmd = new SQLiteCommand("DELETE FROM Exams WHERE Id = @Id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@Id", examId);
                    cmd.ExecuteNonQuery();
                }
                Logger.Info($"Удалён экзамен ID: {examId}");
                if (showInfo) MessageBox.Show("Экзамен удалён");
            }
            catch (Exception ex)
            {
                Logger.Error($"DeleteExam {examId}", ex);
            }
        }

        public void ClearAllExams()
        {
            try
            {
                using (var conn = new SQLiteConnection(_localConnectionString))
                using (var cmd = new SQLiteCommand("DELETE FROM Exams", conn))
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                Logger.Warning("Выполнена полная очистка расписания");
            }
            catch (Exception ex) { Logger.Error("ClearAllExams", ex); }
        }

        // ====================== ВСПОМОГАТЕЛЬНЫЕ ======================
        private string GetTeacherName(int id) => GetNameFromServer("Teachers", "Name", id);
        private string GetSubjectName(int id) => GetNameFromServer("Disciplines", "ShortName9", id);
        private string GetGroupName(int id) => GetNameFromServer("Groups", "Name", id);

        private string GetNameFromServer(string table, string column, int id)
        {
            try
            {
                using (var conn = new SqlConnection(_serverConnectionString))
                using (var cmd = new SqlCommand($"SELECT {column} FROM {table} WHERE Id = @id", conn))
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@id", id);
                    return cmd.ExecuteScalar()?.ToString() ?? "";
                }
            }
            catch { return ""; }
        }

        private string SafeGetString(SQLiteDataReader r, string col)
        {
            try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? "" : r.GetString(i); }
            catch { return ""; }
        }

        private int SafeGetInt32(SQLiteDataReader r, string col)
        {
            try { int i = r.GetOrdinal(col); return r.IsDBNull(i) ? 0 : r.GetInt32(i); }
            catch { return 0; }
        }
        // ====================== СТАРЫЕ МЕТОДЫ ДЛЯ СОВМЕСТИМОСТИ ======================

        /// <summary>
        /// Старый метод для совместимости с MainWindow и DevWindow
        /// </summary>
        public string GetConnectionString()
        {
            return _serverConnectionString;
        }

        /// <summary>
        /// Старый метод — теперь проверяет SQL Server
        /// </summary>
        public void CheckDatabaseStructure()
        {
            CheckServerConnection();
        }

        /// <summary>
        /// Очистка проблемных данных (справочники)
        /// </summary>
        public void CleanProblematicData()
        {
            try
            {
                Logger.Info("Выполняется очистка проблемных данных в справочниках");

                using (var conn = new SqlConnection(_serverConnectionString))
                {
                    conn.Open();

                    string[] queries = {
                        "DELETE FROM Teachers WHERE Name IS NULL OR Name = ''",
                        "DELETE FROM Groups WHERE Name IS NULL OR Name = ''",
                        "DELETE FROM Disciplines WHERE ShortName9 IS NULL OR ShortName9 = ''"
                    };

                    foreach (string q in queries)
                    {
                        using (var cmd = new SqlCommand(q, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                Logger.Info("Очистка проблемных данных завершена");
            }
            catch (Exception ex)
            {
                Logger.Error("CleanProblematicData", ex);
            }
        }
    }
}