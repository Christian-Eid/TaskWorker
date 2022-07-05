using EmailSender.Interfaces;
using EmailSender.Model;
using Serilog;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web;

namespace EmailSender.Services
{
    public class EmailSenderService : IEmailSender
    {
        private readonly ILogger _logger;

        public EmailSenderService(ILogger logger)
        {
            _logger = logger;
        }

        public async Task<bool> SendEmail(MailMessage mail, string source)
        {
            bool ret = true;

            try
            {
                string errorMsg = "";

                #region Get Email Info
                //TODO read from secure Place
                var configuration = ReadJsonFileConfig("/Lookups/EmailConfiguration.json", Directory.GetCurrentDirectory()); // TODO refactor

                string host = configuration.SmtpServer;
                string port = configuration.SmtpPort;
                string ssl = configuration.Ssl;
                string username = configuration.SmtpUsername;
                string pwd = configuration.SmtpPassword;
                string enableSystemEmails = configuration.EnableEmails;
                if (enableSystemEmails == null || !enableSystemEmails.Equals("Y"))
                    return false;
                #endregion

                if (string.IsNullOrWhiteSpace(host) || string.IsNullOrWhiteSpace(port) || string.IsNullOrWhiteSpace(username)
                    || string.IsNullOrWhiteSpace(pwd) ||
                    (string.IsNullOrWhiteSpace(ssl) || (ssl != "Yes") && (ssl != "No")))
                {
                    _logger.Error($"Error Finding Configuration DATA!");
                    return false;
                }

                SmtpClient usrsmtpcheck = new SmtpClient
                {
                    Host = host,
                    Port = int.Parse(port),
                    EnableSsl = ssl.Equals("Yes"),
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(username, pwd)
                };

                try
                {
                    ServicePointManager.ServerCertificateValidationCallback =
                        delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                        { return true; };

                    usrsmtpcheck.Send(mail);
                }
                catch (Exception exp)
                {
                    _logger.Error($"Email Not Delivered!! - {exp}");
                    errorMsg = exp.Message;
                    ret = false;
                }
            }
            catch (Exception exp)
            {
                _logger.Error($"Email Not Delivered! {source} - {mail.Subject} - {exp}");
                ret = false;
            }

            return ret;
        }

        private EmailConfigurationModel ReadJsonFileConfig(string fileName, string filePath) //todo refactor with below method
        {
            EmailConfigurationModel emailConfiguration = new EmailConfigurationModel();
            using (StreamReader r = new StreamReader(filePath + fileName))
            {
                string json = r.ReadToEnd();
                emailConfiguration = JsonConvert.DeserializeObject<EmailConfigurationModel>(json);
            }
            return emailConfiguration;
        }
        private EmailModel ReadJsonFile(string fileName, string filePath) {
            EmailModel emailModel = new EmailModel();
            using (StreamReader r = new StreamReader(filePath + fileName))
            {
                string json = r.ReadToEnd();
                emailModel = JsonConvert.DeserializeObject<EmailModel>(json);
            }
            return emailModel;
        }

        public MailMessage FillMessage(string From, List<string> To, List<string> Bcc, string CatchBouncedEmails, string EmailTemplate, List<string> bodyAttribute, List<string> attachmentPathList)
        {
            MailMessage mail = new MailMessage();

            EmailModel emailModel = ReadJsonFile("/Lookups/EmailTemplate.json", Directory.GetCurrentDirectory()); //TODO refactor

            if (emailModel == null) 
            {
                _logger.Error("Json File ERROR in Reading!!");
                throw new Exception("Error Reading Json");
            }

            if (!string.IsNullOrWhiteSpace(emailModel.From))
            {
                mail.From = new MailAddress(emailModel.From);
            }
            else 
            {
                mail.From = new MailAddress(From);
            }
            
            foreach (var item in To)
            {
                mail.To.Add(item);
            }
            foreach (var item in Bcc)
            {
                mail.Bcc.Add(item);
            }

            if (!string.IsNullOrWhiteSpace(CatchBouncedEmails))
            {
                mail.ReplyToList.Add(new MailAddress(CatchBouncedEmails, "reply-to"));
            }
            else 
            {
                mail.ReplyToList.Add(new MailAddress(emailModel.ReplyTo, "reply-to"));
            }

            if (!string.IsNullOrWhiteSpace(EmailTemplate))
            {
                mail.Subject = emailModel.Subject;
                mail.Body = HttpUtility.HtmlDecode(emailModel.Body);

                int count = 0;
                foreach (var attr in bodyAttribute)
                {
                    mail.Body = mail.Body.Replace("{{" + count + "}}", attr);
                    count++;
                }
                if (count > 0)
                {
                    mail.IsBodyHtml = true;
                }
            }

            foreach (var attachment in attachmentPathList) 
            {
                mail.Attachments.Add(new Attachment(attachment));
            }

            return mail;
        }
    }
}
