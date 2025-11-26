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

namespace ExamScheduleApp.View
{
    /// <summary>
    /// Логика взаимодействия для SubjectDialog.xaml
    /// </summary>
    public partial class SubjectDialog : Window
    {
        public Subject Subject { get; private set; }
        public SubjectDialog()
        {
            InitializeComponent();
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Введите название дисциплины!");
                return;
            }

            Subject = new Subject
            {
                FullName = txtName.Text,
                ShortName12 = "",
                ShortName9 = "",
                ShortName5 = txtCode.Text // Используем поле Code как ShortName5
            };

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
