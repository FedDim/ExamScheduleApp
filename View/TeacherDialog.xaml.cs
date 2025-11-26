using ExamScheduleApp.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExamScheduleApp
{
    /// <summary>
    /// Логика взаимодействия для TeacherDialog.xaml
    /// </summary>
    public partial class TeacherDialog : Window
    {
        public Teacher Teacher { get; private set; }
        public TeacherDialog()
        {
            InitializeComponent();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFullName.Text))
            {
                MessageBox.Show("Введите ФИО преподавателя!");
                return;
            }
            Teacher = new Teacher
            {
                Name = txtFullName.Text,
                Classroom = "", // Значение по умолчанию
                AcademicBuilding = 0 // Значение по умолчанию
            };

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
