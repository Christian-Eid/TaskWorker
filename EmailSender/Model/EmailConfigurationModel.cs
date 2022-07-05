namespace EmailSender.Model
{
    public class EmailConfigurationModel
    {
        public string SmtpServer { get; set; }
        public string SmtpPort { get; set; }
        public string Ssl { get; set; }
        public string SmtpUsername { get; set; }
        public string SmtpPassword { get; set; }
        public string EnableEmails { get; set; }
    }
}
