using ClosedXML.Excel;
using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

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
        private readonly SimpleDatabaseHelper _dbHelper;

        public List<int> ExamsToDelete { get; private set; } = new List<int>();

        public ScheduleTableWindow(ObservableCollection<ExamSchedule> exams, SimpleDatabaseHelper dbhelper)
        {
            InitializeComponent();
            InitializeColumnMapping();

            _originalColumnOrder = new List<DataGridColumn>(ExamsDataGrid.Columns);
            _exams = exams;
            _dbHelper = dbhelper;
            InitializeDataGrid(_exams);

            SortColumnComboBox.SelectedIndex = 0;

            ExamsDataGrid.CellEditEnding += ExamsDataGrid_CellEditEnding;
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

                if (MessageBox.Show($"Вы действительно хотите удалить экзамен/консультацию?\n" +
                $"Преподаватели: {examToDelete.Teacher1Name}, {examToDelete.Teacher2Name}\n" +
                $"Дисциплина: {examToDelete.SubjectName}\n" +
                $"Группа: {examToDelete.GroupName}",
                "Удаление данных", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
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

        private async void GenerateWordButton_Click(object sender, RoutedEventArgs e)
        {
            Button generateButton = sender as Button;

            if (generateButton != null)
            {
                // Сохраняем оригинальный текст кнопки
                string originalText = generateButton.Content.ToString();

                // Блокируем кнопку и меняем текст
                generateButton.IsEnabled = false;
                generateButton.Content = "Формирование...";

                // Можно добавить курсор ожидания
                this.Cursor = Cursors.Wait;

                try
                {
                    MessageBox.Show("Документы начали формироваться. Это может занять некоторое время...",
                                  "Формирование документов",
                                  MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);

                    var saveFileDialog = new SaveFileDialog
                    {
                        FileName = "Выберите папку сохранения",
                        Filter = "Все файлы | *.*",
                        CheckFileExists = false,
                        CheckPathExists = true
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        string selectedPath = Path.GetDirectoryName(saveFileDialog.FileName);

                        if (!string.IsNullOrEmpty(selectedPath))
                        {
                            await Task.Run(() =>
                            {
                                WordHelper wordHelper = new WordHelper(_exams);
                                wordHelper.CreateAllDocuments(selectedPath);
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании документов: {ex.Message}",
                                  "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    // Восстанавливаем исходное состояние
                    generateButton.IsEnabled = true;
                    generateButton.Content = originalText;
                    this.Cursor = Cursors.Arrow;
                }
            }
        }
        /// <summary>
        /// Срабатывает после окончания редактирования ячейки
        /// </summary>
        private void ExamsDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit && e.Row.DataContext is ExamSchedule exam)
            {
                try
                {
                    // Обновляем запись в базе данных
                    UpdateExamInDatabase(exam);

                    Logger.Info($"Обновлён экзамен ID {exam.Id} в таблице");
                }
                catch (Exception ex)
                {
                    Logger.Error($"Ошибка при обновлении экзамена ID {exam.Id}", ex);
                    MessageBox.Show($"Ошибка сохранения изменений: {ex.Message}",
                                  "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateExamInDatabase(ExamSchedule exam)
        {
            if (exam == null || exam.Id <= 0) return;

            try
            {
                using (var conn = new SQLiteConnection(_dbHelper.GetConnectionString())) // используем старый метод
                {
                    conn.Open();
                    string query = @"
                UPDATE Exams 
                SET Teacher1Id = @t1, 
                    Teacher2Id = @t2, 
                    SubjectId = @s, 
                    GroupId = @g,
                    Classroom = @c,
                    Department = @d,
                    ExamDate = @date,
                    ExamTime = @time,
                    ExamType = @type
                WHERE Id = @id";

                    using (var cmd = new SQLiteCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@id", exam.Id);
                        cmd.Parameters.AddWithValue("@t1", exam.Teacher1Id);
                        cmd.Parameters.AddWithValue("@t2", (object)exam.Teacher2Id ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@s", exam.SubjectId);
                        cmd.Parameters.AddWithValue("@g", exam.GroupId);
                        cmd.Parameters.AddWithValue("@c", exam.Classroom ?? "");
                        cmd.Parameters.AddWithValue("@d", exam.DepartmentName ?? "");
                        cmd.Parameters.AddWithValue("@date", exam.ExamDate);
                        cmd.Parameters.AddWithValue("@time", exam.ExamTime);
                        cmd.Parameters.AddWithValue("@type", exam.ExamType);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("UpdateExamInDatabase", ex);
                throw;
            }
        }
    }
}
