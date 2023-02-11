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
using Microsoft.Win32;
using Window = System.Windows.Window;

namespace ExcelReader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public FileExcelReader ExcelReaderObj;
        public MainWindow()
        {
            InitializeComponent();
            ExcelReaderObj = new FileExcelReader();
        }

        public void DisplayMessage(string message)
        {
            Logger.Text += message;
        }

        private void BtnFile_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelReaderObj.FileExcelPath = openFileDialog.FileName;
                FilePathTextBlock.Text = ExcelReaderObj.FileExcelPath;
            }
        }

        private void BtnOpenFileWord_OnClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelReaderObj.FileWordTemplate = openFileDialog.FileName;
                FilePathTextBlockWord.Text = ExcelReaderObj.FileWordTemplate;
            }
        }

        private void BtnStart_OnClick(object sender, RoutedEventArgs e)
        {
            ExcelReaderObj.Notify += DisplayMessage;
            string targetDir = ExcelReaderObj.CreateResultDirectory();
            Logger.Text = targetDir;
            ExcelReaderObj.ReadExcelFile();
            MessageBox.Show($"Done! Find results in {ExcelReaderObj.targetDir}");
        }
    }
}