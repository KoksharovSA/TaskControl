using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        internal Dictionary<string, Grid> ColGrid = new Dictionary<string, Grid>();
        internal Collection<Detail> ColDetail = new Collection<Detail>();
        public MainWindow()
        {
            InitializeComponent();
            ColGrid = LoadTask();
            TreeViewTask.Items.Add("Все");
            foreach (var item in ColGrid.Select(x=> new FileInfo(x.Key).Directory.Name).Distinct().OrderBy(x=>x))
            {
                TreeViewTask.Items.Add(item.Substring(1));
            }
        }

        internal void InfoDetail(IEnumerable<Detail> detail, string dir) 
        {
            InfoDetailWrapPanel.Children.Clear();
            TextBlock textDir = new TextBlock() { Text = dir + "\n" + detail.Select(x=>x.MaterialDetail).FirstOrDefault() + " " + detail.Select(x => x.ThicknessMaterialDetail).FirstOrDefault() + "\n(Осталось разложить)", TextWrapping = TextWrapping.Wrap, FontWeight = FontWeights.UltraBold };
            InfoDetailWrapPanel.Children.Add(textDir);
            foreach (var item in detail)
            {                                
                TextBlock text = new TextBlock() { Text = item.ToString(), TextWrapping = TextWrapping.Wrap };
                InfoDetailWrapPanel.Children.Add(text);
            }
        }

        public Dictionary<string, Grid> LoadTask() 
        {
            Dictionary<string, Grid> dic = new Dictionary<string, Grid>();
            string[] dir = File.ReadAllLines("Dir.txt");
            foreach (var item in dir)
            {
                if (item != "" && item != null)
                {
                    //TreeViewTask.Items.Add(new TreeViewItem() { Header = new FileInfo(item).Directory.Name });
                    Grid grid = new Grid();
                    grid.Margin = new Thickness(5, 5, 5, 5);

                    Rectangle rectangle = new Rectangle();

                    rectangle.StrokeThickness = 2;
                    rectangle.Stroke = Brushes.Green;
                    rectangle.RadiusX = 5;
                    rectangle.RadiusY = 5;

                    StackPanel panel = new StackPanel();
                    panel.Orientation = Orientation.Vertical;
                    panel.HorizontalAlignment = HorizontalAlignment.Center;
                    panel.Width = 200;
                    panel.Margin = new Thickness(5, 5, 5, 5);

                    panel.Children.Add(new TextBlock() { Text = new FileInfo(item).Name + "\n", FontWeight = FontWeights.UltraBold, TextWrapping = TextWrapping.Wrap });
                    ColDetail = ExcelData.ExcelDataLoad(item, 4, new int[] { 1, 2, 12, 13, 35 });
                    foreach (var item1 in ColDetail.Select(x => x.MaterialDetail).Distinct())
                    {
                        foreach (var item2 in ColDetail.Where(x => x.MaterialDetail == item1).OrderBy(x => x.ThicknessMaterialDetail).Select(x => x.ThicknessMaterialDetail).Distinct())
                        {
                            int before = 0;
                            int after = 0;
                            foreach (var item3 in ColDetail.Where(x => x.MaterialDetail == item1 && x.ThicknessMaterialDetail == item2))
                            {
                                before += Convert.ToInt32(item3.QuantityDetail);
                                after += Convert.ToInt32(item3.QuantityDetailNecessary);
                            }
                            if (before != 0)
                            {
                                TextBlock textBlock = new TextBlock() { Text = item1 + "  " + item2 + "  (" + (before - after) + " из " + before + ")", TextWrapping = TextWrapping.Wrap };
                                textBlock.MouseLeftButtonUp += (object sender, MouseButtonEventArgs e) =>
                                {
                                    InfoDetail(ColDetail.Where(x => x.MaterialDetail == item1 && x.ThicknessMaterialDetail == item2 && x.QuantityDetailNecessary != "0"), new FileInfo(item).Name);
                                };
                                panel.Children.Add(textBlock);
                                ProgressBar progressBar = new ProgressBar() { Value = 100 - (after * 100 / before), Height = 20, Width = 200 };
                                panel.Children.Add(progressBar);
                            }
                            else
                            {
                                TextBlock textBlock = new TextBlock() { Text = item1 + "  " + item2 + "  (" + (before - after) + " из " + before + ")", TextWrapping = TextWrapping.Wrap };
                                textBlock.MouseLeftButtonUp += (object sender, MouseButtonEventArgs e) =>
                                {
                                    InfoDetailWrapPanel.Children.Clear();
                                };
                                panel.Children.Add(textBlock);
                                ProgressBar progressBar = new ProgressBar() { Value = 100, Height = 20, Width = 200 };
                                panel.Children.Add(progressBar);
                            }
                        }
                    }
                    grid.Children.Add(rectangle);
                    grid.Children.Add(panel);
                    dic.Add(item,grid);                    
                }               
            }
            return dic;
        }

        private void AddTask_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls";
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            foreach (var item in openFileDialog.FileNames)
            {
                File.AppendAllText("Dir.txt", Environment.NewLine + item);
            }                      
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TreeViewTask_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TaskWrapPanel.Children.Clear();
            if (TreeViewTask.SelectedValue.ToString() == "Все")
            {
                foreach (var item in ColGrid)
                {
                    TaskWrapPanel.Children.Add(item.Value);
                }
            }
            else
            {
                foreach (var item in ColGrid.Where(x => x.Key.Contains(TreeViewTask.SelectedValue.ToString())))
                {
                    TaskWrapPanel.Children.Add(item.Value);
                }
            }
        }

        private void UpdateTask_Click(object sender, RoutedEventArgs e)
        {
            ColGrid = LoadTask();
            TaskWrapPanel.Children.Clear();
        }        
    }
}
