using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;
using static MaterialDesignThemes.Wpf.Theme;

namespace pract9WordExelRed
{
    /// <summary>
    /// Логика взаимодействия для ExelCreateOkno.xaml
    /// </summary>
    public partial class ExelCreateOkno : Window
    {
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        DataTable dataTable1 = new DataTable();

        public ExelCreateOkno()
        {
            InitializeComponent();
            datygridy.ItemsSource = dataTable1.DefaultView;
        }

        private void SaveBtm_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                string selectedPath = dialog.FileName;

                var dataTable = datygridy.ItemsSource as DataView;
                Workbook wb = new Workbook();
                wb.Worksheets.Clear();
                Worksheet sheet = wb.Worksheets.Add("Лист 1");
                sheet.InsertDataView(dataTable, true, 1, 1);

                string filePath = selectedPath + "/" + NameTbox.Text + ".xlsx";

                wb.SaveToFile(filePath, FileFormat.Version2016); ;

                MessageBox.Show("Данные сохранены в файл Excel. По этому пути: " + filePath);
            }
        }

        private void AddColumnBtm_Click(object sender, RoutedEventArgs e)
        {
            DataColumn newColumn = new DataColumn();
            newColumn.ColumnName = NazvanieTbox.Text;
            dataTable1.Columns.Add(newColumn);
            datygridy.ItemsSource = null;
            datygridy.ItemsSource = dataTable1.DefaultView;
        }
    }
}