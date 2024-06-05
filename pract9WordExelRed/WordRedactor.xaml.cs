using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Doc;
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
    /// Логика взаимодействия для WordRedactor.xaml
    /// </summary>
    public partial class WordRedactor : Window
    {
        public string path1;
        public WordRedactor(string path)
        {
            InitializeComponent();

            path1 = path;
            if (path1 != null)
            {
                Document doc = new Document();
                doc.LoadFromFile(path);
                doc.SaveToFile("govno.rtf", FileFormat.Rtf);
                TextRange range = new TextRange(rt.Document.ContentStart, rt.Document.ContentEnd);
                FileStream fStrem = new FileStream("govno.rtf", FileMode.OpenOrCreate);
                range.Load(fStrem, DataFormats.Rtf);
                fStrem.Close();
            }
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TextRange range = new TextRange(rt.Document.ContentStart, rt.Document.ContentEnd);
                FileStream fStrem = new FileStream("govno.rtf", FileMode.Create);
                range.Save(fStrem, DataFormats.Rtf);
                fStrem.Close();

                Document d = new Document();
                d.LoadFromFile("govno.rtf");
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (path1!=null)
                {
                    d.SaveToFile(path1, FileFormat.Docx);
                }
                else
                {
                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.IsFolderPicker = true;
                    CommonFileDialogResult result = dialog.ShowDialog();
                    if (result == CommonFileDialogResult.Ok)
                    {
                        string selectedPath = dialog.FileName;
                        path1 = selectedPath + "\\" + name_file.Text + ".docx";
                        d.SaveToFile(path1, FileFormat.Docx);

                        MessageBox.Show("Данные сохранены в файл Excel. По этому пути: " + path1);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Сначала заполните текстовое поле");
            }
        }

        private void send_Click(object sender, RoutedEventArgs e)
        {
            TextRange range = new TextRange(rt.Document.ContentStart, rt.Document.ContentEnd);
            OtpravkaOkno sendFile = new OtpravkaOkno(path1);
            sendFile.Show();
            Close();
        }
    }
}
