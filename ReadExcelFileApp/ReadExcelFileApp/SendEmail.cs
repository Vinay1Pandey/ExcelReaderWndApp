using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadExcelFileApp
{
    public class SendEmail
    {
        public static void Email(int i, List<string> EmailID)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
                mail.From = new MailAddress("vinay.pandey@gridinfocom.com");
                mail.Subject = "Reminder Request";

                switch (i)
                {
                    case 0:
                        mail.Body = "EM0";
                        break;
                    case 1:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "Your telephone bill numbered NNNN dated dd/mm/yy for an amount of ₹ xx,xx,xxx.xx " +
                            "(Rupees aaa bbb ccc ddd and paise eee only) is due for payment since [due date]. Please pay " +
                            "up this outstanding amount at an early date, to help us in providing you continued better " +
                            "service.Thank you for your prompt action." +
                            "Regards,";
                        break;
                    case 2:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM2";
                        break;
                    case 3:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM3";
                        break;
                    case 4:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM4";
                        break;
                    case 5:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM5";
                        break;
                    case 6:
                        foreach (var item in EmailID)
                        {
                            mail.To.Add(item);
                        }
                        mail.Body = "EM6";
                        break;
                }
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("vinay.pandey@gridinfocom.com", "jj");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
                MessageBox.Show("mail Send");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}
