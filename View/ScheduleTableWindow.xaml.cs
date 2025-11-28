using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExamScheduleApp.View
{
    /// <summary>
    /// Логика взаимодействия для ScheduleTableWindow.xaml
    /// </summary>
    public partial class ScheduleTableWindow : Window
    {
        private ListCollectionView _collectionView;
        private string _currentSortColumn = "ExamDate";
        private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;
        private Dictionary<string, DataGridColumn> _columnMapping;
        private List<DataGridColumn> _originalColumnOrder;
        private ObservableCollection<ExamSchedule> _exams;

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
                { "FirstTeacher", FirstTeacherColumn },
                { "SecondTeacher", SecondTeacherColumn},
                { "ExamDate", DateColumn},
                { "Subject", SubjectColumn},
                { "Group", GroupColumn},
                { "ExamTime", TimeColumn},
                { "Classroom", ClassroomColumn},
                { "ExamType", ExamTypeColumn}
            };

        }

        private void InitializeDataGrid(ObservableCollection<ExamSchedule> exams)
        {
            _collectionView = new ListCollectionView(exams);
            ExamsDataGrid.ItemsSource = _collectionView;

            ApplySorting("ExamDate", ListSortDirection.Ascending);
        }

        private void SortColumnComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplySorting((SortColumnComboBox.SelectedItem as ComboBoxItem).Tag.ToString(), _currentSortDirection);
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

        private void GenerateWordButton_Click(object sender, RoutedEventArgs e)
        {
            WordHelper wordHelper = new WordHelper(_exams);
            wordHelper.CreateDocument();
        }
    }
}
