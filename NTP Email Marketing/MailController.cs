using System;
using System.Collections;
using System.IO;
using Microsoft.VisualBasic;
using System.ComponentModel;
using System.Reflection;
using System.Net.Mail;

namespace NTPEmailMarketing
{
    public class MailController
    {
        public static long Send(MailAccount oMailAccount, string To, string Subject, string Body)
        {
            return Send(oMailAccount, StringToArray(To), null, null, Subject, Body, null);
        }

        public static long Send(MailAccount oMailAccount, string[] To, string Subject, string Body)
        {
            return Send(oMailAccount, To, null, null, Subject, Body, null);
        }

        public static long Send(MailAccount oMailAccount, string To, string Cc, string Bcc, string Subject, string Body, string Attachment)
        {
            return Send(oMailAccount, StringToArray(To), StringToArray(Cc), StringToArray(Bcc), Subject, Body, StringToArray(Attachment));
        }

        /// <summary>
        /// Advanced send
        /// </summary>
        /// <param name="From"></param>
        /// <param name="To"></param>
        /// <param name="CC"></param>
        /// <param name="Bcc"></param>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        /// <param name="Attachments"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long Send(MailAccount oMailAccount, string[] To, string[] CC, string[] Bcc, string Subject, string Body, string[] Attachments, bool Delayed = false)
        {
            try
            {
                //Dim oMailMessage As System.Net.Mail.MailMessage
                Mail oMail = null;

                oMail = new Mail();
                //oMailMessage = New System.Net.Mail.MailMessage()

                // Add the from address
                oMail.From = oMailAccount.Username;
                oMail.FromName = oMailAccount.DisplayName;
                //oMailMessage.Sender = New System.Net.Mail.MailAddress(sFrom)

                if ((To != null))
                {
                    foreach (string sTo in To)
                    {
                        oMail.To.Add(sTo);
                    }
                }

                if ((CC != null))
                {
                    foreach (string sCc in CC)
                    {
                        oMail.Cc.Add(sCc);
                    }
                }

                if ((Bcc != null))
                {
                    foreach (string sBcc in Bcc)
                    {
                        oMail.Bcc.Add(sBcc);
                    }
                }

                oMail.Subject = Subject;
                oMail.Body = Body;
                if ((Attachments != null))
                {
                    foreach (string sAttachment in Attachments)
                    {
                        oMail.Attachments.Add(sAttachment);
                    }
                }

                //if (Delayed)
                    //return SaveDelayedEmail(ref oMail);

                return Send(oMailAccount, oMail);
            }
            catch (Exception ex)
            {
                Information.Err();
                return -1;
            }
        }

        public static MailStatus GetMailStatusList(int MailID)
        {
            MailStatus mailstatus = new MailStatus();

            string[] sFiles;
            string folder;

            mailstatus.OutboxList = new ArrayList();
            mailstatus.SendtList = new ArrayList();
            mailstatus.FailedList = new ArrayList();

            folder = GroupPoster.RootFolder + "\\Mail\\Outbox\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Sendt = sFiles.Length;
                foreach (string file in sFiles)
                {
                    mailstatus.OutboxList.Add(file.Replace(folder, string.Empty).Replace(".dnmail", string.Empty));
                }
                mailstatus.Outbox = sFiles.Length;
            }

            folder = GroupPoster.RootFolder + "\\Mail\\Sendt\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Sendt = sFiles.Length;
                foreach (string file in sFiles)
                {
                    mailstatus.SendtList.Add(file.Replace(folder, string.Empty).Replace(".dnmail", string.Empty));
                }
            }

            folder = GroupPoster.RootFolder + "\\Mail\\Failed\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Failed = sFiles.Length;
                foreach (string file in sFiles)
                {
                    mailstatus.FailedList.Add(file.Replace(folder, string.Empty).Replace(".dnmail", string.Empty));
                }
            }

            return mailstatus;
        }

        public static MailStatus GetMailStatus(int MailID)
        {
            MailStatus mailstatus = new MailStatus();

            string[] sFiles;
            string folder;

            folder = GroupPoster.RootFolder + "\\Mail\\Outbox\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Outbox = sFiles.Length;
            }

            folder = GroupPoster.RootFolder + "\\Mail\\Sendt\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Sendt = sFiles.Length;
            }

            folder = GroupPoster.RootFolder + "\\Mail\\Failed\\" + MailID + "\\";
            if (System.IO.Directory.Exists(folder))
            {
                sFiles = System.IO.Directory.GetFiles(folder);
                mailstatus.Failed = sFiles.Length;
            }

            return mailstatus;
        }

        public class MailStatus
        {
            public int Outbox;
            public int Sendt;
            public int Failed;
            public ArrayList OutboxList;
            public ArrayList SendtList;
            public ArrayList FailedList;
        }
               
        /// <summary>
        /// Advanced send
        /// </summary>
        /// <param name="From"></param>
        /// <param name="To"></param>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static void SendError(string From, string To, string Subject, string Body)
        {
            try
            {
                //Dim oMailMessage As System.Net.Mail.MailMessage
                Mail oMail = null;
                System.Net.Mail.SmtpClient oSmtpClient = null;
                MailAccount oMailAccount = new MailAccount().GetDefaultMailAccount();

                oMail = new Mail();

                // Add the from address
                oMail.From = From;

                oMail.To.Add(To);
                oMail.Subject = Subject;
                oMail.Body = Body;                

                // Sette opp SMTP klient
                oSmtpClient = new System.Net.Mail.SmtpClient(oMailAccount.SMTPHostname, oMailAccount.SMTPPort);
                oSmtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                oSmtpClient.EnableSsl = oMailAccount.SMTPSSL;
                oSmtpClient.Host = oMailAccount.SMTPHostname;

                // Brukernavn og passord, hvis påkrevd
                if (oMailAccount.Username != string.Empty)
                    oSmtpClient.Credentials = new System.Net.NetworkCredential(oMailAccount.Username, oMailAccount.Password);

                // Sende
                oSmtpClient.Send(oMail.ToMailMessage());
            }
            catch (Exception exc)
            {
            }
        }

        [Browsable(false)]
        [EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
        public static string SendInternal(MailAccount oMailAccount, Mail oMail)
        {
            System.Net.Mail.SmtpClient oSmtpClient = null;
            string sMailSMTP = string.Empty;
            int iMailPort = 0;
            string sMailUsername = string.Empty;
            string sMailPassword = string.Empty;
            bool bMailSecure = false;

            sMailSMTP = oMailAccount.SMTPHostname;
            iMailPort = oMailAccount.SMTPPort;
            sMailUsername = oMailAccount.Username;
            sMailPassword = oMailAccount.Password;
            bMailSecure = oMailAccount.SMTPSSL;               

            oSmtpClient = new System.Net.Mail.SmtpClient(sMailSMTP, iMailPort);
            oSmtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            oSmtpClient.EnableSsl = bMailSecure;
            oSmtpClient.Host = sMailSMTP;

            if (sMailUsername != string.Empty) oSmtpClient.Credentials = new System.Net.NetworkCredential(sMailUsername, sMailPassword);

            oSmtpClient.Send(oMail.ToMailMessage());

            return string.Empty;
        }

        /// <summary>
        /// Sender data
        /// </summary>
        /// <param name="oMail"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long Send(MailAccount oMailAccount, Mail oMail, bool Delayed = false)
        {
            // Send the email
            SendInternal(oMailAccount, oMail);
            //if (string.IsNullOrEmpty(SendInternal(oMail)))
            //{
            //    return SaveSendtEmail(ref oMail);
            //}

            // Failed
            return -1;
        }
        
        /// <summary>
        /// Validates an email address
        /// </summary>
        /// <param name="inputEmail"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool isEmail(string inputEmail)
        {
            inputEmail = inputEmail + string.Empty;
            string strRegex = "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$";
            System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(strRegex);
            if ((re.IsMatch(inputEmail)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        
        private static string[] StringToArray(string sStr)
        {
            if (sStr == null)
                return null;
            string[] sArray = new string[1];
            sArray[0] = sStr;
            return sArray;
        }

        private static long SaveSendtEmail(ref Mail oMail)
        {
            try
            {
                long MessageID = 0;
                string sFile = null;
                string sTo = string.Empty;
                string sFolder = string.Empty;

                if (oMail.To.Count > 0)
                    sTo = Convert.ToString(oMail.To[0]);

                MessageID = Convert.ToInt64(Conversion.Val(Strings.Format(System.DateTime.Now, "yyMMddHHmmss") + Math.Round(VBMath.Rnd() * 99, 0)));
                sFile = System.Text.RegularExpressions.Regex.Replace(oMail.Subject, "[^A-Za-z0-9\\u4E00-\\u9FFF ]", string.Empty);
                sTo = System.Text.RegularExpressions.Regex.Replace(sTo, "[^A-Za-z0-9\\.\\@\\u4E00-\\u9FFF ]", string.Empty);

                if (oMail.MailID > 0)
                {
                    sFolder = GroupPoster.RootFolder+"\\Mail\\Sendt\\" + oMail.MailID + "\\";
                    if (!System.IO.Directory.Exists(sFolder)) System.IO.Directory.CreateDirectory(sFolder);
                }
                else
                {
                    sFolder = GroupPoster.RootFolder + "\\Mail\\Sendt\\";
                }

                sFile = sFolder + MessageID + " " + sTo + " " + sFile + ".dnmail";

                oMail.Sent = System.DateTime.Now;
                //System.IO.File.WriteAllText(sFile, Destinet.Utilities.Serialize(oMail, true));

                return MessageID;
            }
            catch (Exception ex)
            {
                Information.Err();
                return -1;
            }
        }

        private static void Err(ref Exception ex)
        {
            string sFileName = null;
            StreamWriter oStreamWriter = null;
            string sError = null;

            sError = Constants.vbCrLf + "Type:    Mail error" + Constants.vbCrLf + "Date:    " + System.DateTime.Now + Constants.vbCrLf + "Message: " + ex.Message.Replace(Constants.vbCrLf, " ") + Constants.vbCrLf + "Stack:   " + ex.StackTrace.Replace(Constants.vbCrLf, " ") + Constants.vbCrLf;

            // Lagre XMl i fil
            sFileName = Strings.Format(DateTime.Now, "yyyyMMdd HHmmss") + ".dnmailerror";
            oStreamWriter = new StreamWriter(sFileName, true);
            oStreamWriter.Write(sError);
            oStreamWriter.Close();
        }
    }
}
