using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Windows;

namespace ExamScheduleApp
{
    public partial class EditTeachersWindow : Window
    {
        private ObservableCollection<Teacher> _teachers;
        private SimpleDatabaseHelper _dbHelper;

        public EditTeachersWindow()
        {
            InitializeComponent();
            _dbHelper = new SimpleDatabaseHelper();
            LoadTeachers();
        }

        private void LoadTeachers()
        {
            try
            {
                var teachers = _dbHelper.GetTeachers();
                _teachers = new ObservableCollection<Teacher>(teachers);
                TeachersDataGrid.ItemsSource = _teachers;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки преподавателей: {ex.Message}");
            }
        }

        private void DeleteTeacher_Click(object sender, RoutedEventArgs e)
        {
            var teacher = (Teacher)((FrameworkElement)sender).DataContext;

            var result = MessageBox.Show(
                $"Вы действительно хотите удалить преподавателя {teacher.Name}?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    _dbHelper.DeleteTeacher(teacher.Id);
                    _teachers.Remove(teacher);
                    MessageBox.Show("Преподаватель удален!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления преподавателя: {ex.Message}");
                }
            }
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (var teacher in _teachers)
                {
                    _dbHelper.UpdateTeacher(teacher);
                }
                MessageBox.Show("Изменения сохранены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения изменений: {ex.Message}");
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}