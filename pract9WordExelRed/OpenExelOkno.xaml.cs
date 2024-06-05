using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
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
    /// Логика взаимодействия для OpenExelOkno.xaml
    /// </summary>
    public partial class OpenExelOkno : Window
    {
        private string path;
        public OpenExelOkno(string path)
        {
            InitializeComponent();
            this.path = path;
            Workbook wb = new Workbook();
            wb.LoadFromFile(path);
            Worksheet sheet = wb.Worksheets[0];
            CellRange locatedRange = sheet.AllocatedRange;

            var dataTable = sheet.ExportDataTable(locatedRange, true);

            datygridy.ItemsSource = dataTable.DefaultView;
        }

        private void AddColumnBtm_Click(object sender, RoutedEventArgs e)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(path);
            Worksheet sheet = wb.Worksheets[0];
            CellRange locatedRange = sheet.AllocatedRange;

            var dataTable = sheet.ExportDataTable(locatedRange, true);
            DataColumn newColumn = new DataColumn();
            newColumn.ColumnName = NazvanieTbox.Text;
            dataTable.Columns.Add(newColumn);
            datygridy.ItemsSource = dataTable.DefaultView;
        }

        private void SaveBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
        }

        private void datygridy_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void SendBtm_Click(object sender, RoutedEventArgs e)
        {
            Save();
            OtpravkaOkno otpravkaOkno = new OtpravkaOkno(path);
            otpravkaOkno.ShowDialog();
        }

        private void Save()
        {
            var dataTable = datygridy.ItemsSource as DataView;

            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet sheet = wb.Worksheets.Add("Лист 1");
            sheet.InsertDataView(dataTable, true, 1, 1);

            wb.SaveToFile(path, FileFormat.Version2016); ;
        }
    }
}
