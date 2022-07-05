using System.Collections.Generic;
using System.Net.Mail;
using System.Threading.Tasks;

namespace EmailSender.Interfaces
{
    public interface IEmailSender
    {
        Task<bool> SendEmail(MailMessage mail, string source);
        MailMessage FillMessage(string From, List<string> To, List<string> Bcc, string CatchBouncedEmails, string EmailTemplate, List<string> bodyAttribute, List<string> attachmentPathList);
    }
}
