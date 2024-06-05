using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Pdf.Exporting.XPS.Schema;
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

namespace pract9WordExelRed
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExelCreateOkno exelCreateOkno = new ExelCreateOkno();
            exelCreateOkno.Show();
            Close();
        }

        private void OpenExelBtm_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            CommonFileDialogResult result = dialog.ShowDialog();
            if(result == CommonFileDialogResult.Ok)
            {
                if (dialog.FileName.Contains(".xlsx"))
                {
                    string path = dialog.FileName;
                    OpenExelOkno openExelOkno = new OpenExelOkno(path); 
                    openExelOkno.Show();
                    Close();
                }
                else
                {
                    MessageBox.Show("Это не эксель");
                }
            }
        }

        private void OpenWord_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                if (dialog.FileName.Contains(".docx"))
                {
                    string path = dialog.FileName;
                    WordRedactor wordRedactor = new WordRedactor(path);
                    wordRedactor.Show();
                    Close();
                }
                else
                {
                    MessageBox.Show("Это не ворд");
                }
            }
        }

        private void CreateWord_Click(object sender, RoutedEventArgs e)
        {
            WordRedactor wordRedactor = new WordRedactor(null);
            wordRedactor.Show();
            Close();
        }
    }
}
