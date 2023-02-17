using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using TransformInsureJToMyReport.ViewModel;

namespace TransformInsureJToMyReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Workbook (*.xlsx)|*.xlsx";
            dialog.Title = "Chọn file trích xuất trực tiếp từ InsureJ";

            if (dialog.ShowDialog() == true)
            {
                if (!SourceFiles.Items.Contains(dialog.FileName))
                {
                    SourceFiles.Items.Add(dialog.FileName);
                }
            }
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (SourceFiles.SelectedIndex != -1)
            {
                SourceFiles.Items.Remove(SourceFiles.SelectedItem);
            }
        }
    }
}
