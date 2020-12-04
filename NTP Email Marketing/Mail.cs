using System;
using System.Collections;
using System.Net.Mail;
using Microsoft.VisualBasic;

namespace NTPEmailMarketing
{
    public class Mail
    {
        public int MailID = 0;
        public string From = string.Empty;
        public string FromName = string.Empty;
        public ArrayList To = new ArrayList();
        public ArrayList Cc = new ArrayList();
        public ArrayList Bcc = new ArrayList();
        public string Subject = string.Empty;
        public string Body = string.Empty;
        public string PlainBody = string.Empty;
        public ArrayList Attachments = new ArrayList();
        public ArrayList AttachedImages = new ArrayList();
        public System.DateTime Submitted;
        public System.DateTime Sent;
        public string Error = string.Empty;
        public bool IsBodyHtml = true;

        /// <summary>
        /// Konvertere til MailMessage
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public System.Net.Mail.MailMessage ToMailMessage()
        {
            System.Net.Mail.MailMessage oMailMessage = null;
            string sMediaType = string.Empty;
            System.Net.Mail.LinkedResource oLinkedResource = null;
            int iImageCount = 0;

            oMailMessage = new System.Net.Mail.MailMessage();

            string sFrom = this.From;
            if (sFrom.Length == 0) sFrom = "tommynguyen24612@gmail.com";

            string sFromName = this.FromName;
            if (sFromName.Length == 0) sFromName = "Săn Khuyến Mãi";

            oMailMessage.From = new System.Net.Mail.MailAddress(sFrom, sFromName);
            oMailMessage.Sender = new System.Net.Mail.MailAddress(sFrom, sFromName);

            foreach (string sMailAddress in this.To)
            {
                oMailMessage.To.Add(sMailAddress);
            }
            foreach (string sMailAddress in this.Cc)
            {
                oMailMessage.CC.Add(sMailAddress);
            }
            foreach (string sMailAddress in this.Bcc)
            {
                oMailMessage.Bcc.Add(sMailAddress);
            }

            oMailMessage.Subject = this.Subject.Replace(Constants.vbCr, "").Replace(Constants.vbLf, "");
            oMailMessage.Body = this.Body;

            oMailMessage.IsBodyHtml = this.IsBodyHtml;

            foreach (string sAttachment in Attachments)
            {
                oMailMessage.Attachments.Add(new System.Net.Mail.Attachment(sAttachment));
            }

            if (this.AttachedImages.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(PlainBody))
                {
                    System.Net.Mail.AlternateView plainTextView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(PlainBody, null, "text/plain");
                    oMailMessage.AlternateViews.Add(plainTextView);
                }

                System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(Body, null, "text/html");

                foreach (string sImage in this.AttachedImages)
                {
                    if (System.IO.File.Exists(sImage))
                    {

                        if (sImage.EndsWith(".jpg") | sImage.EndsWith(".jpeg")) sMediaType = System.Net.Mime.MediaTypeNames.Image.Jpeg;
                        if (sImage.EndsWith(".png")) sMediaType = "binary/png";
                        if (sImage.EndsWith(".gif")) sMediaType = System.Net.Mime.MediaTypeNames.Image.Gif;

                        oLinkedResource = new System.Net.Mail.LinkedResource(sImage, sMediaType);
                        oLinkedResource.ContentId = "Image" + iImageCount;
                        oLinkedResource.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
                        oLinkedResource.ContentId = "Image" + iImageCount;

                        htmlView.LinkedResources.Add(oLinkedResource);
                        iImageCount += 1;
                    }
                }

                //oMailMessage.AlternateViews.Add(plainTextView);
                oMailMessage.AlternateViews.Add(htmlView);
            }
            else
            {
                oMailMessage.Body = this.Body;
            }

            return oMailMessage;
        }

        private string html2plaintext(string html)
        {
            string plainhtml = string.Empty;

            try
            {
                // Breaks into linebreaks and remove windows linebreaks
                plainhtml = html.Replace("<br />", "\n");
                plainhtml = plainhtml.Replace("\r", "");

                // Remove tags
                plainhtml = System.Text.RegularExpressions.Regex.Replace(plainhtml, "<.*?>", " ");

                string[] plainhtmlsplit = plainhtml.Split('\n');

                System.Text.StringBuilder newplainhtml = new System.Text.StringBuilder();

                int lbcount = 0;
                foreach (string line in plainhtmlsplit)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        newplainhtml.Append(line.Trim() + "\n");
                        lbcount = 0;
                    }
                    else
                    {
                        if (lbcount == 0)
                        {
                            lbcount++;
                            newplainhtml.Append("\n");
                        }
                    }
                }

                // Remove duplicate spaces
                System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None;
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"[ ]{2,}", options);
                plainhtml = regex.Replace(newplainhtml.ToString(), @" ");

            }
            catch (System.Exception ex)
            {
                return string.Empty;
            }

            return plainhtml;
        }
    }
}
