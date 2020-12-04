using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace NTPEmailMarketing
{
    public class Handler
    {
        public static void SendEmail(MailAccount oMailAccount, string subject, string to, string body, bool isPostCard)
        {
            StringBuilder sbOperation = null;
            try
            {
                Mail oMail = new Mail();
                oMail.Subject = subject;
                oMail.From = oMailAccount.Username;
                oMail.FromName = oMailAccount.DisplayName;
                oMail.To.Add(to);
                oMail.Body = body;
                oMail.IsBodyHtml = true;
                MailController.SendInternal(oMailAccount, oMail);
                if (isPostCard)
                {
                    EmailMarketing.Insert(to, subject, DateTime.Now.ToString("dd/MM/yyyy"), 0);
                }
                else
                {
                    EmailMarketing.Insert(to, subject, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), 0);
                }                
                // Log finish to send email
                sbOperation = new StringBuilder();
                sbOperation.AppendLine("Finish posting to group!");
                sbOperation.AppendLine("Data inserted into database");
                Log(GroupPoster.logFilePath, sbOperation.ToString());
                Thread.Sleep(3 * 1000);
            }
            catch (Exception ex)
            {
                // Log error
                sbOperation = new StringBuilder();
                sbOperation.AppendLine("Error when posting!");
                sbOperation.AppendLine("Error: " + ex.Message);
                Log(GroupPoster.logFilePath, sbOperation.ToString());
                if (ex.Message == "Failure sending mail." || ex.Message.Contains("Mailbox unavailable. The server response was: 5.4.5 Daily user sending quota exceeded"))
                {
                    //EmailMarketing.Insert(to, subject, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), 0);
                    //Log(NTPEmailMarketing.rejectedFilePath, to);
                    InsertDelay(GroupPoster.delayFilePath, 24);
                    Thread.Sleep(24 * 60 * 60 * 1000);
                }
            }            
        }

        public static void SendEmail(MailAccount oMailAccount, string subject, string []to, string body)
        {
            StringBuilder sbOperation = null;
            try
            {
                MailController.Send(oMailAccount, to, subject, body);
                foreach (var email in to)
                {
                    EmailMarketing.Insert(email.ToLower(), subject, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), 0);  
                }
                // Log finish to send email
                sbOperation = new StringBuilder();
                sbOperation.AppendLine("Finish posting to group!");
                sbOperation.AppendLine("Data inserted into database");
                Log(GroupPoster.logFilePath, sbOperation.ToString());
                InsertDelay(GroupPoster.delayFilePath, 4);
                Thread.Sleep(4 * 60 * 60 * 1000);
            }
            catch (Exception ex)
            {
                // Log error
                sbOperation = new StringBuilder();
                sbOperation.AppendLine("Error when posting!");
                sbOperation.AppendLine("Error: " + ex.Message);
                Log(GroupPoster.logFilePath, sbOperation.ToString());
                if (ex.Message == "Failure sending mail." || ex.Message.Contains("Mailbox unavailable. The server response was: 5.4.5 Daily user sending quota exceeded"))
                {
                    //EmailMarketing.Insert(to, subject, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), 0);
                    //Log(NTPEmailMarketing.rejectedFilePath, to);
                    InsertDelay(GroupPoster.delayFilePath, 24);
                    Thread.Sleep(24 * 60 * 60 * 1000);
                }
            }
        }

        // Check if Internet is connected?
        public static bool NetworkIsAvailable()
        {
            var all = NetworkInterface.GetAllNetworkInterfaces();
            foreach (var item in all)
            {
                if (item.NetworkInterfaceType == NetworkInterfaceType.Loopback || item.NetworkInterfaceType == NetworkInterfaceType.Tunnel)
                    continue;
                //if (item.Name.ToLower().Contains("virtual") || item.Description.ToLower().Contains("virtual"))
                //    continue; //Exclude virtual networks set up by VMWare and others
                if (item.OperationalStatus == OperationalStatus.Up)
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsAvailableNetworkActive()
        {
            // only recognizes changes related to Internet adapters
            if (NetworkInterface.GetIsNetworkAvailable())
            {
                // however, this will include all adapters -- filter by opstatus and activity
                NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
                return (from face in interfaces
                        where face.OperationalStatus == OperationalStatus.Up
                        where (face.NetworkInterfaceType != NetworkInterfaceType.Tunnel) && (face.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                        select face.GetIPv4Statistics()).Any(statistics => (statistics.BytesReceived > 0) && (statistics.BytesSent > 0));
            }
            return false;
        }

        private static void KillProcess()
        {
            var processList = Process.GetProcessesByName("WINWORD");
            foreach (var item in processList)
            {
                item.Kill();
            }
        }

        // Convert file from word to html
        public static void ConvertFile(string postFolder)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object oMiss = System.Reflection.Missing.Value;
            object oTrue = true;
            object oFalse = false;
            object fltDocFormat = 10;

            try
            {
                string[] postFileList = GetFiles(postFolder, GroupPoster.PATTERN_WORD);
                foreach (var postFile in postFileList)
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    object file1 = postFile;
                    object file2 = postFile.Replace(".docx", ".html");
                    if (File.Exists(file2.ToString()))
                    {
                        wordApp.Quit(ref oMiss, ref oMiss, ref oMiss);
                        KillProcess();
                        continue;
                    }
                    Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(
                                ref file1, ref oFalse, ref oTrue, ref oFalse,
                                ref oMiss, ref oMiss, ref oTrue, ref oMiss,
                                ref oMiss, ref oMiss, ref oMiss, ref oFalse,
                                ref oFalse, Microsoft.Office.Interop.Word.WdDocumentDirection.wdLeftToRight,
                                ref oTrue, ref oMiss);
                    wordApp.Visible = false;
                    //doc.Fields.Update();
                    doc.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

                    doc.SaveAs2(ref file2, ref fltDocFormat, ref oMiss, ref oMiss,
                                ref oMiss, ref oMiss, ref oMiss, ref oTrue,
                                ref oMiss, ref oMiss, ref oMiss, ref oMiss,
                                ref oTrue, ref oMiss, ref oMiss, ref oMiss,
                                ref oMiss);

                    doc.Close(ref oMiss, ref oMiss, ref oMiss);
                    wordApp.Quit(ref oMiss, ref oMiss, ref oMiss);
                    KillProcess();
                }
            }
            catch (Exception ex)
            {
                wordApp.Quit();
                KillProcess();
            }
        }

        // Read content of file word
        public static string ReadContent(string path)
        {
            string result = "";
            Microsoft.Office.Interop.Word.Application wordApp = null;
            object file = path;
            object oMiss = System.Reflection.Missing.Value;
            object oTrue = true;
            object oFalse = false;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(
                    ref file, ref oFalse, ref oTrue, ref oFalse,
                    ref oMiss, ref oMiss, ref oTrue, ref oMiss,
                    ref oMiss, ref oMiss, ref oMiss, ref oFalse,
                    ref oFalse, Microsoft.Office.Interop.Word.WdDocumentDirection.wdLeftToRight,
                    ref oTrue, ref oMiss);

                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();

                System.Windows.Forms.IDataObject data = System.Windows.Forms.Clipboard.GetDataObject();

                result = data.GetData(System.Windows.Forms.DataFormats.Html).ToString();
                //int iStart = result.IndexOf("<html");
                result = result.Substring(result.IndexOf("<html"));
                doc.Close(ref oMiss, ref oMiss, ref oMiss);
                wordApp.Quit(ref oMiss, ref oMiss, ref oMiss);
                KillProcess();

            }
            catch (Exception ex)
            {
                wordApp.Quit();
                KillProcess();
                return null;
            }

            return result;
        }

        //using a regular expression, find all of the href or urls in the content of the page 
        public static bool isValidEmail(string email)
        {
            //regular expression 
            string pattern = @"^[a-z0-9._-]+@(?:yahoo\.com|yahoo\.com\.vn|gmail\.com|hotmail\.com)$";

            //Set up regex object 
            Regex reg = new Regex(pattern, RegexOptions.IgnoreCase);

            return reg.IsMatch(email);
        }

        public static string GetValidEmail(string sEmail)
        {
            return sEmail.Replace(",", "")
                         .Replace(";", "")
                         .Replace("#", "")
                         .Replace("!", "")
                         .Replace("%", "")
                         .Replace("$", "")
                         .Replace("&", "")
                         .Replace("+", "")
                         .Replace("=", "")
                         .Replace("^", "")
                         .ToLower();
        }

        public static string[] GetFiles(string SourceFolder, string pattern)
        {
            //string[] sDirs = null;
            string[] sFiles = null;

            if (!System.IO.Directory.Exists(SourceFolder))
                return null;

            // Subfolders
            //sDirs = System.IO.Directory.GetDirectories(SourceFolder);
            //foreach (string sDir in sDirs)
            //{
            //    GetFiles(sDir, pattern);
            //}
            sFiles = System.IO.Directory.GetFiles(SourceFolder, pattern, SearchOption.TopDirectoryOnly).Where(f => !f.Contains("$")).ToArray();
            return sFiles;
        }
                
        // Save log of operations
        public static void Log(string path, string operation)
        {
            try
            {                
                using (TextWriter tw = new StreamWriter(path, true))
                {
                    tw.WriteLine(operation);
                }
            }
            catch (Exception)
            {
            }
        }

        // Insert delay time of sending email
        public static void InsertDelay(string path, int delay)
        {
            try
            {                
                using (TextWriter tw = new StreamWriter(path))
                {
                    tw.Write(delay);
                }
            }
            catch (Exception)
            {                
            }
        }

        // Get delay time of sending email
        public static int GetDelay(string path)
        {
            int ret = 0;
            try
            {
                if (!File.Exists(path))
                {
                    InsertDelay(path, 0);
                }                
                using (TextReader tr = new StreamReader(path))
                {
                    ret = int.Parse(tr.ReadLine().Trim());
                    tr.Close();
                }  
            }
            catch (Exception)
            {
                return 0;
            }                       
            return ret;
        }

        // Check if database has rejected emails, and then update rejected
        public static void UpdateRejected(string path)
        {            
            try
            {
                var emailsList = File.ReadAllLines(path)
                                    .Where(arg => !string.IsNullOrWhiteSpace(arg))
                                    .Distinct()
                                    .ToList();
                foreach (var email in emailsList)
                {
                    if (!EmailMarketing.IsRejected(email))
                    {
                        EmailMarketing.Update(email, 1);
                    }
                }
            }
            catch (Exception)
            {                
            }            
        }

        // Url: https://365.vtc.vn/v2/tin-tuc/tin-khuyen-mai/79/p=1
        // Regex: @"text-tabs-tintuc.*?<a href=\""[viettel|mobifone|vinaphone].*?-dd-MM-.*?\"">(.*?)</a>"
        // Return title and datetime

        const string sUrl1 = "https://365.vtc.vn/v2/tin-tuc/tin-khuyen-mai/79/p=1";
        const string sUrl2 = "https://thanhtoanonline.vn/tin-tuc/3/Thong-tin-khuyen-mai.html";
        public static string[,] GetPromotionNews(int index)
        {            
            string[,] newsList = null;            
            try
            {
                bool isInternetConnected = true;
                do
                {
                    isInternetConnected = NetworkIsAvailable();
                } while (!isInternetConnected);

                HttpSession httpDownload = new HttpSession(); 
                string sDownloadContent = "", strRegex = "";
                string sTitle = "", sLink = ""; 
                DateTime dTime = DateTime.Now;
                Regex rxGetResult = null;
                MatchCollection matchCollection = null;
                int count = 0;

                string day = DateTime.Now.Day >= 10 ? DateTime.Now.Day.ToString() : ("0" + DateTime.Now.Day);
                string month = DateTime.Now.Month >= 10 ? DateTime.Now.Month.ToString() : ("0" + DateTime.Now.Month);
                string year = DateTime.Now.Year.ToString();

                if (index == 1)
                {
                    sDownloadContent = httpDownload.GetMethodDownload(sUrl1, true, false, false, false);
                    sDownloadContent = System.Web.HttpUtility.HtmlDecode(sDownloadContent);                    
                    strRegex = @"text-tabs-tintuc.*?<a href=\""viettel|mobifone|vinaphone.*?-" + day + "-" + month + @"-.*?\"">(.*?)</a>";

                    rxGetResult = new Regex(strRegex, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    matchCollection = rxGetResult.Matches(sDownloadContent);
                    count = matchCollection.Count;
                    newsList = new string[count, 2];
                    for (int i = 0; i < count; i++)
                    {
                        sTitle = matchCollection[i].Groups[1].Value.Trim();
                        sTitle = sTitle.Replace("NGÀY VÀNG", "50% giá trị thẻ nạp duy nhất ngày");
                        sTitle = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(sTitle.ToLower());
                        newsList[i, 0] = dTime.ToString("dd/MM/yyyy");
                        newsList[i, 1] = sTitle;
                    }
                }
                else
                {
                    sDownloadContent = httpDownload.GetMethodDownload(sUrl2, true, false, false, false);
                    sDownloadContent = System.Web.HttpUtility.HtmlDecode(sDownloadContent);
                    strRegex = @"<a class=\""g-title\"" href=\""(.*?)\"">.*?" + day + "/" + month + "/" + @".*?Viettel|Mobifone|Vinaphone.*?</a>";

                    rxGetResult = new Regex(strRegex, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                    matchCollection = rxGetResult.Matches(sDownloadContent);
                    count = matchCollection.Count;
                    newsList = new string[count, 2];
                    for (int i = 0; i < count; i++)
                    {
                        sLink = matchCollection[i].Groups[1].Value;
                        sDownloadContent = httpDownload.GetMethodDownload(sLink, true, false, false, false);
                        sDownloadContent = System.Web.HttpUtility.HtmlDecode(sDownloadContent);
                        strRegex = @"<font size=\""2\"" color=\""#ff0000\"">(.*?)</font>";

                        rxGetResult = new Regex(strRegex, RegexOptions.Singleline | RegexOptions.IgnoreCase);
                        Match match = rxGetResult.Match(sDownloadContent);
                        if (match.Success)
                        {
                            sTitle = match.Groups[1].Value.Trim();
                        }
                        else
                        {
                            sTitle = "";
                        }
                        
                        newsList[i, 0] = dTime.ToString("dd/MM/yyyy");
                        newsList[i, 1] = sTitle;   
                    }
                }                
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                return null;
            }
            return newsList;
        }
    }
}
