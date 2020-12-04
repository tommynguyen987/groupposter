using System.Net;
using System.Net.Mail;

namespace NTPEmailMarketing
{
    public class MailAccount
    {
        public string DisplayName = "Săn Khuyến Mãi";
        public string SMTPHostname = "smtp.gmail.com";
        public string Username = "tommynguyen24612@gmail.com";
        public string Password = "phatIT0687!@#123";
        public int SMTPPort = 587;
        public bool SMTPSSL = true;

        public SmtpClient GetSmtpClient()
        {
            var client = new SmtpClient(SMTPHostname, SMTPPort)
            {
                DeliveryMethod = SmtpDeliveryMethod.Network,
                EnableSsl = SMTPSSL,
                Host = SMTPHostname,
                Credentials = new NetworkCredential(Username, Password)
            };

            return client;
        }

        public MailAccount GetDefaultMailAccount()
        {
            MailAccount oMailAccount = new MailAccount();
            var _with1 = oMailAccount;
            _with1.DisplayName = DisplayName;
            _with1.Username = Username;
            _with1.Password = Password;
            _with1.SMTPHostname = SMTPHostname;
            _with1.SMTPPort = SMTPPort;
            _with1.SMTPSSL = true;
            return oMailAccount;
        }
    }
}
