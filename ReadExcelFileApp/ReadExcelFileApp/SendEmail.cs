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
        public static void Email(int i, string EmailID,string LastPayDate, string ListBill,string Reminder,string DueDate,string CustName,string Gadgets)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.IsBodyHtml = true;
                SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
                mail.From = new MailAddress("vinay.pandey@gridinfocom.com");
                mail.Subject = "Reminder Request";
                switch (i)
                {
                    case 0:
                        
                            mail.To.Add(EmailID);
                        
                        mail.Body = "EM0";
                        break;
                    case 1:
                        
                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p> Your telephone dated <b> " + LastPayDate + " </b></p>" +
                                        " <p>for an amount of <b>₹" + ListBill + " </b></p>" +
                                        " <p>is due for payment since <b>" + LastPayDate + "</b>.Please pay up this outstanding </p>" +
                                        " <p> amount at an early date, to help us in providing you continued better service.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6></p> ";                        
                        break;
                    case 2:
                        
                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p> Further to our earlier request, we bring to your attention that your telephone </p>" +
                                        "<p> dated <b>  " + LastPayDate + "</b></p>" +
                                        " <p>for an amount of <b>₹ " + ListBill + " </b></p>" +
                                        " <p>is due for payment since <b>" + LastPayDate + " </b>.Please pay up this outstanding </p>" +
                                        " <p> amount at an early date, to help us in providing you continued better service.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6></p> ";
                        
                        break;
                    case 3:
                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p> Further to our earlier request, we would again bring to your attention that your telephone </p>" +
                                        "<p> dated <b> " + LastPayDate + " </b></p>" +
                                        " <p>for an amount of <b>₹" + ListBill+ " </b></p>" +
                                        " <p>is due for payment since <b>" + LastPayDate + " </b>.Please pay up this outstanding </p>" +
                                        " <p> amount at an early date, to help us in providing you continued better service.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6 ></p> ";
                        
                        break;
                    case 4:

                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p>This is our <b>" + Reminder + "th</b> request to pay your outstanding due</p>" +
                                        " <p> Further to our earlier request, we would again bring to your attention that your telephone </p>" +
                                        "<p> dated <b> " + LastPayDate + " </b></p>" +
                                        " <p>for an amount of <b>₹ " + ListBill + " </b></p>" +
                                        " <p>is due for payment since <b>" + LastPayDate + " </b>.Please pay up this outstanding </p>" +
                                        " <p> amount at an early date, to help us in providing you continued better service.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6></p> ";
                        
                        break;
                    case 5:

                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p>This is our <b>" + Reminder + "th</b> request to pay your outstanding due</p>" +
                                        " <p> Further to our earlier request, we would again bring to your attention that your telephone </p>" +
                                        "<p> dated <b>" + LastPayDate + "</b></p>" +
                                        " <p>for an amount of <b>₹ " + ListBill + " </b></p>" +
                                        " <p>is due for payment since <b>" + LastPayDate + " </b>.Please pay up this outstanding </p>" +
                                        " <p> amount at an early date, so as to avoid disconnection of your <b>"+Gadgets+"</b>.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6></p> ";
                        
                        break;
                    case 6:

                            mail.To.Add(EmailID);

                            mail.Body = " <div align='justify' style='font - family: Times New Roman; font - size: 15px; '>" +
                                        " <strong> Dear " + CustName + "</strong>,&nbsp;" +
                                        " <p>This is our <b>" + Reminder + "th</b> request to pay your outstanding due</p>" +
                                        " <p> Further to our earlier request, we would again bring to your attention that your telephone</p>" +
                                        "<p> dated <b> " + LastPayDate + " </b></p>" +
                                        " <p>for an amount of <b>₹ " + ListBill + " </b></p>" +
                                        " <p>is due for payment since < b >" + LastPayDate + " </ b >.Please pay up this outstanding </p>" +
                                        " <p> amount immediately, so as to avoid disconnection of your <b>" + Gadgets + "</b>.</p>" +
                                        " <p><strong> Thank you for your prompt action.</strong></p>" +
                                        " <p><b> Regards,</b></p>" +
                                        " <p>BSNL Team</p>"+
                                        " </div>" +
                                        " <p><h6> This email is system generated, please do not respond to this email.</h6></p> ";
                        
                        break;
                }
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("vinay.pandey@gridinfocom.com", "9899938450@vinay");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
                mail.To.Clear();
                //MessageBox.Show("mail Send");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}
