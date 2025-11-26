using System;
using System.Collections.ObjectModel;
using System.Windows;
using ExamScheduleApp.Model;

namespace ExamScheduleApp
{
    public partial class EditGroupsWindow : Window
    {
        private ObservableCollection<Group> _groups;
        private SimpleDatabaseHelper _dbHelper;

        public EditGroupsWindow()
        {
            InitializeComponent();
            _dbHelper = new SimpleDatabaseHelper();
            LoadGroups();
        }

        private void LoadGroups()
        {
            try
            {
                var groups = _dbHelper.GetGroups();
                _groups = new ObservableCollection<Group>(groups);
                GroupsDataGrid.ItemsSource = _groups;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}");
            }
        }

        private void DeleteGroup_Click(object sender, RoutedEventArgs e)
        {
            var group = (Group)((FrameworkElement)sender).DataContext;

            var result = MessageBox.Show(
                $"Вы действительно хотите удалить группу {group.Name}?",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    _dbHelper.DeleteGroup(group.Id);
                    _groups.Remove(group);
                    MessageBox.Show("Группа удалена!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка удаления группы: {ex.Message}");
                }
            }
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (var group in _groups)
                {
                    _dbHelper.UpdateGroup(group);
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