using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using ExamScheduleApp.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

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
        private ObservableCollection<ExamSchedule> _exams;
        public SimpleDatabaseHelper dbhelper = new SimpleDatabaseHelper();

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                // Диагностика базы данных
                dbhelper.CheckDatabaseStructure();

                // Обновляем существующие записи
                dbhelper.UpdateExamFields();

                // Очищаем проблемные данные (для SimpleDatabaseHelper.cs)
                dbhelper.CleanProblematicData();

                // Очищаем все экзамены при запуске
                //dbhelper.ClearAllExams();

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

                if (_exams is null)
                {
                    if (MessageBox.Show("Загузить данные с буфера? Иначе они будут очищены", "Загрузка буфера", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    {
                        if (MessageBox.Show("Вы точно уверены, что хотите очистить буфер?", "Очистка буфера", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            dbhelper.ClearAllExams();
                            MessageBox.Show("Буфер очищен", "Информация");
                        }
                    }

                    _exams = new ObservableCollection<ExamSchedule>();
                }

                RefreshExamsData();

                // Убрать после наладки
                //MessageBox.Show($"Загружено: {_teachers.Count} преподавателей, {_subjects.Count} дисциплин, {_groups.Count} групп, {_exams.Count} экзаменов");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных из базы: {ex.Message}");
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
                    DepartmentName = group.Department,
                    ExamDate = dateString,
                    ExamTime = time,
                    ExamType = type
                };

                // Добавление в базу данных
                dbhelper.AddExam(exam);

                // Получаем ID добавленной записи
                var addedExams = dbhelper.GetExamSchedule();
                if (addedExams.Count > 0)
                {
                    exam.Id = addedExams[addedExams.Count - 1].Id;
                }

                // Обновление интерфейса
                _exams.Add(exam);
                //ExamsDataGrid.Items.Refresh();

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

        private void EditTeachers_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.TEACHER);
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditSubjects_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.SUBJECT);
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void EditGroups_Click(object sender, RoutedEventArgs e)
        {
            var window = new EditWindow(DataType.GROUP);
            window.Closed += (s, args) => LoadDataFromDatabase(); // Обновляем данные после закрытия окна
            window.ShowDialog();
        }

        private void ShowScheduleTable_Click(object sender, RoutedEventArgs e)
        {
            var scheduleTableWindow = new ScheduleTableWindow(_exams);
            scheduleTableWindow.Owner = this;
            scheduleTableWindow.Closed += ScheduleTableWindow_Closed;
            scheduleTableWindow.ShowDialog();
        }

        private void ScheduleTableWindow_Closed(object sender, EventArgs e)
        {
            var scheduleTableWindow = sender as ScheduleTableWindow;

            if (scheduleTableWindow?.ExamsToDelete?.Count > 0)
            {
                try
                {
                    // Удаляем экзамены из базы данных
                    DeleteExamsFromDatabase(scheduleTableWindow.ExamsToDelete);

                    // Обновляем данные
                    RefreshExamsData();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении экзаменов: {ex.Message}");
                }
            }
        }

        private void DeleteExamsFromDatabase(List<int> examIds)
        {
            foreach (var examId in examIds)
            {
                dbhelper.DeleteExam(examId);
            }
        }

        private void RefreshExamsData()
        {
            try
            {
                // Обновляем коллекцию экзаменов из базы данных
                var examList = dbhelper.GetExamSchedule();
                _exams.Clear();
                foreach (var exam in examList)
                {
                    _exams.Add(exam);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных: {ex.Message}");
            }
        }
    }
}
