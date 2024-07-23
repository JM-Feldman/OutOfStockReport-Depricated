using System;
using System.Diagnostics;
using System.Net.Mail;
using System.Threading; //using System.Web.Mail;

namespace OOSReport.Classes
{
    /// 
    /// Class to handle emails.    
    /// 
    public class eMail
    {
        #region SendeMail

        public static void SendeMail(string subject, string message, string attachement)
        {
            // Send mail
            SendeMail(CommonSettings.GetSetting("AdminEmail"), "OOS.Report@spar.co.za", subject,
                message, attachement, false);
        }

        #endregion

        public static void Send(MailMessage mailMessage)
        {
            // Queue the task and data.
            ThreadPool.QueueUserWorkItem(threadSendeMail, mailMessage);
        }

        #region threadSendeMail

        private static void threadSendeMail(object mailInfo)
        {
            try
            {
                // Create smtp obj
                var SmtpServer = new SmtpClient(CommonSettings.GetSetting("MailServer"));

                // Send mail
                SmtpServer.Send((MailMessage) mailInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        #endregion

        #region Old Code SendMail

        //public static void SendMail(string subject, string message, string attachement)
        //{
        //    // Send mail
        //    sendMail(CommonSettings.GetSetting("AdminEmail"),"Catman.SparcatII.DataSelection@spar.co.za",subject,
        //        message,
        //        MailFormat.Text, attachement);
        //}
        //#endregion

        //#region sendMail
        //private static void sendMail(string recipient,string sender,string subject,string message, MailFormat mailFormat, string attachement )
        //{
        //    // new mail object
        //    MailInfo mi = new MailInfo(recipient, sender, subject,message, mailFormat, attachement );


        //    // Queue the task and data.
        //    ThreadPool.QueueUserWorkItem(new WaitCallback(threadSendMail), mi);			
        //}
        //#endregion

        //#region threadSendMail
        //private static void threadSendMail(Object mailInfo) 
        //{
        //    MailInfo mailInf = (MailInfo) mailInfo;

        //    // Create new mail Message
        //    System.Web.Mail.MailMessage mailMessage = new System.Web.Mail.MailMessage(); 
        //    mailMessage.To = mailInf._Recipient;  
        //    mailMessage.From = mailInf._Sender;
        //    mailMessage.Subject = mailInf._Subject;
        //    mailMessage.BodyFormat = mailInf._MailFormat;
        //    mailMessage.Body = mailInf._Message;

        //    // Add attachment
        //    if(mailInf._Attachement != null && File.Exists(mailInf._Attachement))
        //    {
        //        MailAttachment newAttachment = new MailAttachment(mailInf._Attachement, MailEncoding.UUEncode);
        //        mailMessage.Attachments.Add(newAttachment);
        //    }

        //    // Create smtp obj
        //    SmtpMail.SmtpServer = ConfigurationSettings.AppSettings["MailServer"];

        //    // Send mail
        //    SmtpMail.Send(mailMessage);
        //}

        #endregion

        #region SendeMail

        public static void SendeMail(string recipients, string sender, string subject, string message,
            string attachements, bool isBodyHtml)
        {
            SendeMail(recipients, sender, subject, message, attachements, isBodyHtml, true);
        }

        public static void SendeMail(string recipients, string sender, string subject, string message,
            string attachements, bool isBodyHtml, bool useThreadPool)
        {
            // new mail object
            var mm = new MailMessage();

            mm.From = new MailAddress(sender);
            mm.Sender = new MailAddress(sender);
            mm.Subject = subject;
            mm.Body = message;
            mm.IsBodyHtml = true;
            mm.Priority = MailPriority.Normal;
            //if (!recipients.ToLower().Contains("rhyno.linde@spar.co.za")) mm.Bcc.Add(new MailAddress("Rhyno.Linde@spar.co.za"));
            if (!Debugger.IsAttached)
            {
                //if (!recipients.ToLower().Contains("catmansupport@spar.co.za")) mm.Bcc.Add(new MailAddress("CatmanSupport@spar.co.za"));
            }

            mm.IsBodyHtml = isBodyHtml;

            recipients = recipients.Replace(",", ";");
            foreach (var to in recipients.Split(';')) mm.To.Add(new MailAddress(to));

            if (attachements != null)
                foreach (var att in attachements.Split(';'))
                    mm.Attachments.Add(new Attachment(att));

            if (useThreadPool)
                // Queue the task and data.
                ThreadPool.QueueUserWorkItem(threadSendeMail, mm);
            else
                threadSendeMail(mm);
        }

        #endregion
    }
}