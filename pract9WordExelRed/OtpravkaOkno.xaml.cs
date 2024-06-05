using ImapX.Collections;
using ImapX.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
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

namespace pract9WordExelRed
{
    /// <summary>
    /// Логика взаимодействия для OtpravkaOkno.xaml
    /// </summary>
    public partial class OtpravkaOkno : Window
    {
        private string path;
        public OtpravkaOkno(string path)
        {
            InitializeComponent();
            this.path = path;
        }

        private void SendBtm_Click(object sender, RoutedEventArgs e)
        {
            int port = 25;
            if (YourMailTbox.Text.Contains("@gmail.com") == true)
            {
                string protokol = "gmail.com";
                Otpravka(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@mail.ru") == true)
            {
                string protokol = "mail.ru";
                port = 587;
                Otpravka(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@rambler.ru")== true)
            {
                string protokol = "rambler.ru";
                Otpravka(protokol, port);
            }
            else if (YourMailTbox.Text.Contains("@yandex.ru") == true)
            {
                string protokol = "yandex.ru";
                Otpravka(protokol, port);
            }
            else
            {
                MessageBox.Show("Введён недоступный протокол");
            }
        }

        private void Otpravka(string protokol, int port)
        {
            MailMessage mailMessage = new MailMessage(YourMailTbox.Text, FriendMailTbox.Text, ThemeTbox.Text, "Я вообще в шоке");
            mailMessage.Attachments.Add(new Attachment(path));
            SmtpClient smtpClient = new SmtpClient("smtp." + protokol, port);
            smtpClient.Credentials = new NetworkCredential(YourMailTbox.Text, PasswordTbox.Text);
            smtpClient.EnableSsl = true;    
            smtpClient.Send(mailMessage);
            MessageBox.Show("Сообщение отправилось!");
            //schemetovyaSa@yandex.ru

            //testwpfc@mail.ru

            //asdfghfdgfdg@gmail.com

            //govno123121@ranbler.ru
            //хз не получилось взять
        }
    }
}
