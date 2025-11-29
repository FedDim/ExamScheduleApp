using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Windows;

namespace ExamScheduleApp
{
    public partial class EditSubjectsWindow : Window
    {
        private ObservableCollection<Subject> _subjects;
        private SimpleDatabaseHelper _dbHelper;

        public EditSubjectsWindow()
        {
            InitializeComponent();
            _dbHelper = new SimpleDatabaseHelper();
            LoadSubjects();
        }

        private void LoadSubjects()
        {
            try
            {
                var subjects = _dbHelper.GetSubjects();
                _subjects = new ObservableCollection<Subject>(subjects);
                SubjectsDataGrid.ItemsSource = _subjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки дисциплин: {ex.Message}");
            }
        }

        private void DeleteSubject_Click(object sender, RoutedEventArgs e)
        {
            var subject = (Subject)((FrameworkElement)sender).DataContext;

            var result = MessageBox.Show(
                $"Вы действительно хотите удалить дисциплину {subject.ShortName9}?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    _dbHelper.DeleteSubject(subject.Id);
                    _subjects.Remove(subject);
                    MessageBox.Show("Дисциплина удалена!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления дисциплины: {ex.Message}");
                }
            }
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (var subject in _subjects)
                {
                    _dbHelper.UpdateSubject(subject);
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