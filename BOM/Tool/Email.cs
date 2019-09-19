using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;

namespace BOM.Tool
{
    public static class Email
    {
        //private static string EMAIL_FROM = "KISTJLI@schaeffler.com";
        //private static string EMAIL_FROM = "diazian@schaeffler.com";
        private static readonly string EMAIL_FROM = "SANTOVCT@schaeffler.com";
        public static  void Send(string toEmailAddress, string emailSubject, string emailMessage)
        {

            MailMessage mailMessage = new MailMessage();
            MailAddress from = new MailAddress(EMAIL_FROM);
            SmtpClient smtpClient = new SmtpClient("mail.emea.luk.com");

            mailMessage.From = from;
            mailMessage.Subject = emailSubject;
            mailMessage.Body = emailMessage;
            mailMessage.Priority = MailPriority.High;
            mailMessage.To.Add(new MailAddress(toEmailAddress));
            
            mailMessage.IsBodyHtml = true;
            smtpClient.Send(mailMessage);
            
        }
        

    }
}
