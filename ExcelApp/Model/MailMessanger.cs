using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Collections.ObjectModel;

namespace ExcelApp.Model
{
    class MailMessanger
    {
        ObservableCollection<string> _emailList;
        string _pathExcel;
        string _pathWord;

        public MailMessanger(ObservableCollection<string> emailList, string pathExcel = "c:\\1\\1.xlsx", string pathWord = "c:\\1\\1.docx")
        {
            _emailList = emailList;
            _pathExcel = pathExcel;
            _pathWord = pathWord;
        }
        public void SendMail()
        {
            SmtpClient smtpClient = new SmtpClient("smtp.yandex.ru", 25);
            smtpClient.Credentials = new NetworkCredential("evgeny.yandutov", "evgeny.yandutov1");
            smtpClient.EnableSsl = true;
            foreach (string adress in _emailList)
            {
                MailMessage mailMessage = new MailMessage("evgeny.yandutov@yandex.ru", adress);
                mailMessage.Subject = "Отчёт";
                mailMessage.Body = "Добрый день, отправляю вам отчёт";
                Attachment Excel = new Attachment(_pathExcel);
                Attachment Word = new Attachment(_pathWord);
                mailMessage.Attachments.Add(Excel);
                mailMessage.Attachments.Add(Word);
                smtpClient.Send(mailMessage);
            }
        }
    }
}
