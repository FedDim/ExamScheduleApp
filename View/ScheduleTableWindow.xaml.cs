using ExamScheduleApp.Model;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExamScheduleApp.View
{
    public partial class ScheduleTableWindow : Window
    {
        private ListCollectionView _collectionView;
        private string _currentSortColumn = "Teacher1Name";
        private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
        private Dictionary<string, DataGridColumn> _columnMapping;
        private List<DataGridColumn> _originalColumnOrder;
        private ObservableCollection<ExamSchedule> _exams;
        private ExamSchedule _selectedItem;

        public List<int> ExamsToDelete { get; private set; } = new List<int>();

        public ScheduleTableWindow(ObservableCollection<ExamSchedule> exams)
        {
            InitializeComponent();
            InitializeColumnMapping();

            _originalColumnOrder = new List<DataGridColumn>(ExamsDataGrid.Columns);
            _exams = exams;
            InitializeDataGrid(_exams);

            SortColumnComboBox.SelectedIndex = 0;
        }

        private void InitializeColumnMapping()
        {
            _columnMapping = new Dictionary<string, DataGridColumn>
            {
                { "Teacher1Name", FirstTeacherColumn },
                { "Teacher2Name", SecondTeacherColumn},
                { "ExamDate", DateColumn},
                { "SubjectName", SubjectColumn},
                { "GroupName", GroupColumn},
                { "DepartmentName", DepartmentColumn},
                { "ExamTime", TimeColumn},
                { "Classroom", ClassroomColumn},
                { "ExamType", ExamTypeColumn}
            };

        }

        private void InitializeDataGrid(ObservableCollection<ExamSchedule> exams)
        {
            _collectionView = new ListCollectionView(exams);
            ExamsDataGrid.ItemsSource = _collectionView;

            ApplySorting(_currentSortColumn, ListSortDirection.Ascending);

            ResettingLineSelection();
        }

        private void SortColumnComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedItem = (SortColumnComboBox.SelectedItem as ComboBoxItem).Tag.ToString();

            ApplySorting(selectedItem, _currentSortDirection);
        }

        private void AscendingButton_Click(object sender, RoutedEventArgs e)
        {
            if (SortColumnComboBox.SelectedItem != null)
            {
                ApplySorting(_currentSortColumn, ListSortDirection.Ascending);

                AscendingButton.IsEnabled = false;
                DescendingButton.IsEnabled = true;
            }
        }

        private void DescendingButton_Click(object sender, RoutedEventArgs e)
        {
            if (SortColumnComboBox.SelectedItem != null)
            {
                ApplySorting(_currentSortColumn, ListSortDirection.Descending);

                DescendingButton.IsEnabled = false;
                AscendingButton.IsEnabled = true;
            }
        }

        private void ApplySorting(string sortBy, ListSortDirection direction)
        {
            _currentSortColumn = sortBy;
            _currentSortDirection = direction;

            _collectionView.SortDescriptions.Clear();
            _collectionView.SortDescriptions.Add(new SortDescription(sortBy, direction));
            _collectionView.Refresh();

            MoveSortColumnToFront(sortBy);
        }

        private void MoveSortColumnToFront(string sortTag)
        {
            ExamsDataGrid.Columns.Clear();

            foreach (var column in _originalColumnOrder) ExamsDataGrid.Columns.Add(column);

            if (_columnMapping.TryGetValue(sortTag, out DataGridColumn sortColumn))
            {
                ExamsDataGrid.Columns.Remove(sortColumn);
                ExamsDataGrid.Columns.Insert(0, sortColumn);
            }
        }

        private void ExamsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _selectedItem = (ExamSchedule)ExamsDataGrid.SelectedItem;

            if (_selectedItem != null) DeleteExam.IsEnabled = _selectedItem != null;
        }

        private void DeleteExam_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedItem != null)
            {
                var examToDelete = _selectedItem;

                if (MessageBox.Show("Удалить экзамен/консультацию?", "Удаление данных", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    ExamsToDelete.Add(examToDelete.Id);

                    _exams.Remove(examToDelete);
                    ApplySorting(_currentSortColumn, _currentSortDirection);

                    ResettingLineSelection();
                }
            }
        }

        private void ResettingLineSelection()
        {
            ExamsDataGrid.SelectedItem = null;
            _selectedItem = null;
            DeleteExam.IsEnabled = false;
        }

        //private void GenerateWordButton_Click(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {
        //        var saveFileDialog = new SaveFileDialog
        //        {
        //            FileName = "Выберите папку сохранения",
        //            Filter = "Все файлы | *.*",
        //            CheckFileExists = false,
        //            CheckPathExists = true
        //        };

        //        if (saveFileDialog.ShowDialog() == true)
        //        {
        //            string selectedPath = Path.GetDirectoryName(saveFileDialog.FileName);

        //            if (!string.IsNullOrEmpty(selectedPath))
        //            {
        //                WordHelper wordHelper = new WordHelper(_exams);
        //                wordHelper.CreateAllDocuments(selectedPath);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Ошибка : {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        //    }
        //}
    }
}
