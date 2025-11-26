using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Windows;

public class SimpleDatabaseHelper
{
    private string GetConnectionString()
    {
        string basePath = AppDomain.CurrentDomain.BaseDirectory;
        string dbPath = Path.Combine(basePath, "Data", "ExamScheduleDB.db");
        return $"Data Source={dbPath};Version=3;";
    }

    public void CheckDatabaseStructure()
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();

                // Получаем список всех таблиц
                string tablesQuery = "SELECT name FROM sqlite_master WHERE type='table';";
                using (var command = new SQLiteCommand(tablesQuery, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        List<string> tables = new List<string>();
                        while (reader.Read())
                        {
                            tables.Add(reader.GetString(0));
                        }
                    }
                }

                // Проверяем структуру каждой таблицы
                CheckTableStructure(connection, "Teachers");
                CheckTableStructure(connection, "Groups");
                CheckTableStructure(connection, "Disciplines");
                CheckTableStructure(connection, "Exams");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка проверки структуры БД: {ex.Message}");
        }
    }

    private void CheckTableStructure(SQLiteConnection connection, string tableName)
    {
        try
        {
            string query = $"PRAGMA table_info({tableName});";
            using (var command = new SQLiteCommand(query, connection))
            {
                using (var reader = command.ExecuteReader())
                {
                    List<string> columns = new List<string>();
                    while (reader.Read())
                    {
                        string columnName = reader.GetString(1);
                        string columnType = reader.GetString(2);
                        columns.Add($"{columnName} ({columnType})");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка проверки таблицы {tableName}: {ex.Message}");
        }
    }


    public List<Teacher> GetTeachers()
    {
        var teachers = new List<Teacher>();

        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "SELECT id, name, classroom, academicBuilding FROM Teachers ORDER BY name";

                using (var command = new SQLiteCommand(query, connection))
                {
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
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка загрузки преподавателей: {ex.Message}");
        }

        return teachers;
    }

    public List<Subject> GetSubjects()
    {
        var subjects = new List<Subject>();

        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "SELECT id, fullname, shortName12, shortName9, shortName5 FROM Disciplines ORDER BY fullname";

                using (var command = new SQLiteCommand(query, connection))
                {
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
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка загрузки дисциплин: {ex.Message}");
        }

        return subjects;
    }

    public List<Group> GetGroups()
    {
        var groups = new List<Group>();

        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "SELECT id, name FROM Groups ORDER BY name";

                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            groups.Add(new Group
                            {
                                Id = SafeGetInt32(reader, "id"),
                                Name = SafeGetString(reader, "name")
                            });
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка загрузки групп: {ex.Message}");
        }

        return groups;
    }

    public List<ExamSchedule> GetExamSchedule()
    {
        var exams = new List<ExamSchedule>();

        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();

                string query = "SELECT Id, Teacher1Id, Teacher2Id, SubjectId, GroupId, Classroom FROM Exams";

                using (var command = new SQLiteCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
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
                                ExamDate = "",
                                ExamTime = "",
                                ExamType = ""
                            };

                            // Получаем названия по ID
                            exam.Teacher1Name = GetTeacherName(exam.Teacher1Id);
                            exam.Teacher2Name = GetTeacherName(exam.Teacher2Id);
                            exam.SubjectName = GetSubjectName(exam.SubjectId);
                            exam.GroupName = GetGroupName(exam.GroupId);

                            exams.Add(exam);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка загрузки экзаменов: {ex.Message}");
        }

        return exams;
    }

    private string SafeGetString(SQLiteDataReader reader, string columnName)
    {
        try
        {
            int columnIndex = reader.GetOrdinal(columnName);
            if (!reader.IsDBNull(columnIndex))
                return reader.GetString(columnIndex);
            else
                return string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private int SafeGetInt32(SQLiteDataReader reader, string columnName)
    {
        try
        {
            int columnIndex = reader.GetOrdinal(columnName);
            if (!reader.IsDBNull(columnIndex))
                return reader.GetInt32(columnIndex);
            else
                return 0;
        }
        catch
        {
            return 0;
        }
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
                    var result = command.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }
        catch
        {
            return "";
        }
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
                    var result = command.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }
        catch
        {
            return "";
        }
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
                    var result = command.ExecuteScalar();
                    return result?.ToString() ?? "";
                }
            }
        }
        catch
        {
            return "";
        }
    }

    public void CleanProblematicData()
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();

                // Удаляем записи с NULL значениями в важных полях
                string[] cleanupQueries = {
                "DELETE FROM Teachers WHERE name IS NULL OR classroom IS NULL",
                "DELETE FROM Groups WHERE name IS NULL",
                "DELETE FROM Disciplines WHERE shortName9 IS NULL",
                "DELETE FROM Exams WHERE Teacher1Id IS NULL OR Teacher2Id IS NULL OR SubjectId IS NULL OR GroupId IS NULL OR Classroom IS NULL"
            };

                foreach (string query in cleanupQueries)
                {
                    using (var command = new SQLiteCommand(query, connection))
                    {
                        int affected = command.ExecuteNonQuery();
                        if (affected > 0)
                            MessageBox.Show($"Удалено {affected} проблемных записей");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка очистки данных: {ex.Message}");
        }
    }

    public void ClearAllExams()
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "DELETE FROM Exams";

                using (var command = new SQLiteCommand(query, connection))
                {
                    int rowsDeleted = command.ExecuteNonQuery();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка очистки экзаменов: {ex.Message}");
        }
    }

    public void AddExam(ExamSchedule exam)
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = @"
                    INSERT INTO Exams (Teacher1Id, Teacher2Id, SubjectId, GroupId, Classroom) 
                    VALUES (@Teacher1Id, @Teacher2Id, @SubjectId, @GroupId, @Classroom)";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Teacher1Id", exam.Teacher1Id);
                    command.Parameters.AddWithValue("@Teacher2Id", exam.Teacher2Id);
                    command.Parameters.AddWithValue("@SubjectId", exam.SubjectId);
                    command.Parameters.AddWithValue("@GroupId", exam.GroupId);
                    command.Parameters.AddWithValue("@Classroom", exam.Classroom);

                    command.ExecuteNonQuery();
                }
            }

            MessageBox.Show("Экзамен успешно добавлен в базу данных!");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка добавления экзамена: {ex.Message}");
        }
    }

    public void DeleteExam(int examId)
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "DELETE FROM Exams WHERE Id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Id", examId);
                    int rowsDeleted = command.ExecuteNonQuery();

                    if (rowsDeleted > 0)
                    {
                        MessageBox.Show("Экзамен успешно удален из базы данных");
                    }
                    else
                    {
                        MessageBox.Show("Не удалось найти экзамен для удаления");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка удаления экзамена: {ex.Message}");
            throw;
        }
    }

    public void AddTeacher(Teacher teacher)
    {
        string query = "INSERT INTO Teachers (name, classroom, academicBuilding) VALUES (@Name, @Classroom, @AcademicBuilding)";

        using (var connection = new SQLiteConnection(GetConnectionString()))
        {
            connection.Open();
            using (var command = new SQLiteCommand(query, connection))
            {
                command.Parameters.Add(new SQLiteParameter("@Name", DbType.String) { Value = teacher.Name });
                command.Parameters.Add(new SQLiteParameter("@Classroom", DbType.String) { Value = teacher.Classroom });
                command.Parameters.Add(new SQLiteParameter("@AcademicBuilding", DbType.Int32) { Value = teacher.AcademicBuilding });
                command.ExecuteNonQuery();
            }
        }
    }

    public void AddSubject(Subject subject)
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "INSERT INTO Disciplines (fullname, shortName12, shortName9, shortName5) VALUES (@FullName, @ShortName12, @ShortName9, @ShortName5)";

                using (var command = new SQLiteCommand(query, connection))
                {
                    // Используем shortName9 как основное имя, а остальные генерируем автоматически
                    string shortName9 = subject.ShortName9;

                    command.Parameters.AddWithValue("@FullName", shortName9); // Используем shortName9 как полное имя
                    command.Parameters.AddWithValue("@ShortName12", shortName9.Length > 12 ? shortName9.Substring(0, 12) : shortName9);
                    command.Parameters.AddWithValue("@ShortName9", shortName9);
                    command.Parameters.AddWithValue("@ShortName5", shortName9.Length > 5 ? shortName9.Substring(0, 5) : shortName9);

                    command.ExecuteNonQuery();
                }
            }

            MessageBox.Show($"Дисциплина '{subject.ShortName9}' успешно добавлена!");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка добавления дисциплины: {ex.Message}");
            throw;
        }
    }

    public void AddGroup(Group group)
    {
        string query = "INSERT INTO Groups (name) VALUES (@Name)";

        using (var connection = new SQLiteConnection(GetConnectionString()))
        {
            connection.Open();
            using (var command = new SQLiteCommand(query, connection))
            {
                command.Parameters.Add(new SQLiteParameter("@Name", DbType.String) { Value = group.Name });
                command.ExecuteNonQuery();
            }
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
                command.Parameters.Add(new SQLiteParameter("@Name", DbType.String) { Value = name });
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
                command.Parameters.Add(new SQLiteParameter("@Name", DbType.String) { Value = name });
                var count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
            }
        }
    }

    // Методы для работы с преподавателями
    public void UpdateTeacher(Teacher teacher)
    {
        try
        {
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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка обновления преподавателя: {ex.Message}");
            throw;
        }
    }

    public void DeleteTeacher(int teacherId)
    {
        try
        {
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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка удаления преподавателя: {ex.Message}");
            throw;
        }
    }

    // Методы для работы с дисциплинами
    public void UpdateSubject(Subject subject)
    {
        try
        {
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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка обновления дисциплины: {ex.Message}");
            throw;
        }
    }

    public void DeleteSubject(int subjectId)
    {
        try
        {
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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка удаления дисциплины: {ex.Message}");
            throw;
        }
    }

    // Методы для работы с группами
    public void UpdateGroup(Group group)
    {
        try
        {
            using (var connection = new SQLiteConnection(GetConnectionString()))
            {
                connection.Open();
                string query = "UPDATE Groups SET name = @Name WHERE id = @Id";

                using (var command = new SQLiteCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Name", group.Name);
                    command.Parameters.AddWithValue("@Id", group.Id);

                    command.ExecuteNonQuery();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка обновления группы: {ex.Message}");
            throw;
        }
    }

    public void DeleteGroup(int groupId)
    {
        try
        {
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
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка удаления группы: {ex.Message}");
            throw;
        }
    }
}