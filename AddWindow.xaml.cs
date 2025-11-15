using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Логика взаимодействия для AddWindow.xaml
    /// </summary>
    public partial class AddWindow : Window
    {
        public static string data;
        public AddWindow(string text)
        {
            InitializeComponent();
            AddTextBlock.Text = $"Введите {text}";
        }

        private void AddButtonClick(object sender, RoutedEventArgs e)
        {
            if(AddTextBox.Text != "")
            {
                if (AddTextBlock.Text.Equals("Введите группу"))
                {
                    if (!Regex.IsMatch(AddTextBox.Text.Trim(), @"^[А-ЯЁ]{2}-\d{2}$"))
                    {
                        MessageBox.Show("Введите корректное название группы! Пример: ДВ-45");
                    }
                    else
                    {
                        AddData(AddTextBox.Text);
                    }
                }
                else
                {
                    AddData(AddTextBox.Text);
                }
            }
            else
            {
                MessageBox.Show("Заполните поле!");
            }

        }

        private void AddData(string text)
        {
            data = AddTextBox.Text;
            MessageBox.Show("Данные обновлены!");
            this.Close(); ;
        }
    }
}
