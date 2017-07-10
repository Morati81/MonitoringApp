using System;
using System.Configuration;
using System.Net.Mail;

namespace MonitoringApp
{
    class Mail
    {
        public static bool SendMail(string Oggetto, string body,string allegato="")
        {
            try
            {
                string MailFrom = ConfigurationManager.AppSettings["MailFrom"];
                string MailTo = ConfigurationManager.AppSettings["MailTo"];
                string MailCC = "";
                string MailCCN = "";



                MailMessage oMsg = new MailMessage();
                // Set the message sender
                oMsg.From = new MailAddress(MailFrom, ConfigurationManager.AppSettings["MailFromDisplay"]);

                // The .To property is a generic collection,
                // so we can add as many recipients as we like.
                oMsg.To.Add(MailTo);
                //new MailAddress(
                if (MailCC != "")
                    oMsg.CC.Add(MailCC.Replace(";", ","));

                if (MailCCN != "")
                    oMsg.Bcc.Add(MailCCN.Replace(";", ","));
                if (allegato != "")
                {
                    oMsg.Attachments.Add(new Attachment(allegato));
                }

                // Set the content
                oMsg.Subject = Oggetto;

                oMsg.Body = body;
                oMsg.IsBodyHtml = true;

                SmtpClient oSmtp = new SmtpClient(ConfigurationManager.AppSettings["ServerSMTP"]);
                oSmtp.EnableSsl =bool.Parse(ConfigurationManager.AppSettings["EnableSsl"]);  

                //You can choose several delivery methods. 
                //Here we will use direct network delivery.
                oSmtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //Some SMTP server will require that you first 
                //authenticate against the server.
                System.Net.NetworkCredential oCredential = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["UserSMTP"], ConfigurationManager.AppSettings["PasswordSMTP"]);
                oSmtp.UseDefaultCredentials = false;
                oSmtp.Credentials = oCredential;
                //Let's send it already
                oSmtp.Send(oMsg);



                return true;
            }
            catch(Exception ex)  { throw; }
        }

        
    }
}
