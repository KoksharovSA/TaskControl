using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace TaskСontrol
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void AddTask_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName != "" && openFileDialog.FileName != null)
            {
                TreeViewTask.Items.Add(new TreeViewItem() { Header = new FileInfo(openFileDialog.FileName).Directory.Name });
                foreach (var item in ExcelData.ExcelDataLoad(openFileDialog.FileName, 2, new int[]{ 1,2,12,13})) 
                { 
                    
                }
            }
        }
    }
}
