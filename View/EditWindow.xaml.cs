using ExamScheduleApp.Model;
using ExamScheduleApp.Utilities;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace ExamScheduleApp.View
{
    /// <summary>
    /// Логика взаимодействия для EditWindow.xaml
    /// </summary>
    public partial class EditWindow : Window
    {
        private DataType _dataType = DataType.NULL;
        private SimpleDatabaseHelper _dbHelper;

        private ObservableCollection<Teacher> _teachers;
        private ObservableCollection<Subject> _subjects;
        private ObservableCollection<Group> _groups;

        public EditWindow(DataType dataType)
        {
            InitializeComponent();

            _dataType = dataType;
            _dbHelper = new SimpleDatabaseHelper();

            switch (_dataType)
            {
                case DataType.SUBJECT:
                    LoadSubjects();
                    break;
                case DataType.GROUP:
                    LoadGroups();
                    break;
                case DataType.TEACHER:
                    LoadTeachers();
                    break;
                default:
                    MessageBox.Show("Такой тип отсутствует");
                    break;
            }
        }

        private void LoadGroups()
        {
            try
            {
                var groups = _dbHelper.GetGroups();
                _groups = new ObservableCollection<Group>(groups);

                EditDataGrid.Columns.Clear();

                CreateTextColumn("Название группые", "Name", 200);
                CreateTextColumn("Отделение", "Department", 200);
                CreateActionColumn("Действия", 100, new RoutedEventHandler(DeleteGroup_Click));

                EditDataGrid.ItemsSource = _groups;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}");
            }
        }

        private void LoadSubjects()
        {
            try
            {
                var subjects = _dbHelper.GetSubjects();
                _subjects = new ObservableCollection<Subject>(subjects);

                this.Width = 700;
                this.Height = 400;

                EditDataGrid.Columns.Clear();

                CreateTextColumn("Полное название", "FullName", 200);
                CreateTextColumn("Сокр. (12)", "ShortName12", 100);
                CreateTextColumn("Сокр. (9)", "ShortName9", 100);
                CreateTextColumn("Сокр. (5)", "ShortName5", 80);
                CreateActionColumn("Действия", 100, new RoutedEventHandler(DeleteSubject_Click));

                EditDataGrid.ItemsSource = _subjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки дисциплин: {ex.Message}");
            }
        }

        private void LoadTeachers()
        {
            try
            {
                var teachers = _dbHelper.GetTeachers();
                _teachers = new ObservableCollection<Teacher>(teachers);

                EditDataGrid.Columns.Clear();

                CreateTextColumn("ФИО", "Name", 200);
                CreateTextColumn("Кабинет", "Classroom", 100);
                CreateTextColumn("Корпус", "AcademicBuilding", 80);
                CreateActionColumn("Действия", 100, new RoutedEventHandler(DeleteTeacher_Click));

                EditDataGrid.ItemsSource = _teachers;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки преподавателей: {ex.Message}");
            }
        }

        private void CreateTextColumn(string header, string bindingPath, double width)
        {
            DataGridTextColumn column = new DataGridTextColumn
            {
                Header = header,
                Binding = new Binding(bindingPath)
                {
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                },
                Width = new DataGridLength(width)
            };

            EditDataGrid.Columns.Add(column);
        }

        private void CreateActionColumn(string header, double width, RoutedEventHandler eventHandler)
        {
            DataGridTemplateColumn column = new DataGridTemplateColumn
            {
                Header = header,
                Width = new DataGridLength(width)
            };

            // Создаем DataTemplate
            DataTemplate template = new DataTemplate();

            // Создаем кнопку через FrameworkElementFactory
            FrameworkElementFactory buttonFactory = new FrameworkElementFactory(typeof(Button));
            buttonFactory.SetValue(Button.ContentProperty, "Удалить");

            // Применяем стиль вместо прямого задания цветов
            Style redButtonStyle = Application.Current.FindResource("RedModernButtonStyle") as Style;
            if (redButtonStyle != null)
            {
                buttonFactory.SetValue(Button.StyleProperty, redButtonStyle);
            }
            else
            {
                // Fallback на прямые цвета если стиль не найден
                buttonFactory.SetValue(Button.BackgroundProperty, Brushes.Red);
                buttonFactory.SetValue(Button.ForegroundProperty, Brushes.White);
            }

            buttonFactory.SetValue(Button.MarginProperty, new Thickness(2));
            buttonFactory.SetValue(Button.PaddingProperty, new Thickness(8, 4, 8, 4)); // Сохраняем похожий размер
            buttonFactory.AddHandler(Button.ClickEvent, eventHandler);

            template.VisualTree = buttonFactory;
            column.CellTemplate = template;

            EditDataGrid.Columns.Add(column);
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

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Escape:
                    Close();
                    break;
            }
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            switch (_dataType)
            {
                case DataType.SUBJECT:
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
                    break;
                case DataType.GROUP:
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
                    break;
                case DataType.TEACHER:
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
                    break;
                default:
                    MessageBox.Show("Такой тип отсутствует");
                    break;
            }
        }
    }
}
