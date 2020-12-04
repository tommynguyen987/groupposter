using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net;
using System.Text;
using System.Collections;
using System.Web;
using System.Text.RegularExpressions;
using System.Globalization;

namespace NTPEmailMarketing
{
    public partial class GroupPoster : Form
    {
        public GroupPoster()
        {
            InitializeComponent();
            //StartUpManager.AddApplicationToCurrentUserStartup("Group Poster");
        }
        public static string RootFolder = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
        public static string CategoryFolder = Directory.GetParent(RootFolder).FullName + "\\Categories\\";        
        public static string logFilePath = Application.StartupPath + "\\logs.txt";
        public static string rejectedFilePath = Application.StartupPath + "\\rejected.txt";
        public static string delayFilePath = Application.StartupPath + "\\delay.txt";
               
        public const string PATTERN_ALL = "*.*";
        public const string PATTERN_WORD = "*.docx";
        public const string PATTERN_HTML = "*.html";
        public static string[] Categories = { "SanKhuyenMai", "ThuThuatMayTinh", "ThuGianMoiNgay", "KienThucMoiNgay", "KienThucDanhChoNam", "KienThucDanhChoNu" };
        Thread[] listThreads;
        private bool isRunning;

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;

        [DllImport("user32.dll")]
        static extern bool AnimateWindow(System.IntPtr hWnd, int time, AnimateWindowFlags flags);
        [System.Flags]
        enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x00000001,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        private void WindowsInSystemTray(bool inTray)
        {
            if (inTray)
            {
                this.ShowInTaskbar = false;
                AnimateWindow(this.Handle, 50, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                myNotifyIcon.Visible = true;
                myNotifyIcon.ShowBalloonTip(500);
            }
            else
            {
                this.ShowInTaskbar = true;
                this.WindowState = FormWindowState.Normal;
                AnimateWindow(this.Handle, 700, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_ACTIVATE);
                this.Activate();
                myNotifyIcon.Visible = false;
            }
        }

        private void PostCardGroup()
        {
            string[,] results = Handler.GetPromotionNews(1);
            int len1 = results==null?0:int.Parse(results.GetLength(0).ToString());
            int length = 0;
            if (len1 > 0)
            {
                length = len1;
            }
            else
            {
                results = Handler.GetPromotionNews(2);
                int len2 = results == null ? 0 : int.Parse(results.GetLength(0).ToString());
                if (len2 > 0)
                {
                    length = len2;
                }
            }            

            string to = "dichvunaptienthecao@groups.facebook.com";
            for (int i = 0; i < length; i++)
            {
                string subject = "[" + results[i, 0] + "] THÔNG TIN KHUYẾN MẠI";
                string message = results[i, 1] + Environment.NewLine;
                message += "Mại dô! Mại dô! :)";
                SendGmail(subject, to, message);
                Thread.Sleep(3 * 60 * 1000);
            }
        }

        private void SendGmail(string subject, string to, string body)
        {
            MailAccount oMailAccount = new MailAccount();
            oMailAccount.DisplayName = "Dịch Vụ Nạp Tiền - Thẻ Cào";  
            // Send email
            if (!EmailMarketing.IsExistTitle(body, DateTime.Now.ToString("dd/MM/yyyy")))
            {
                Handler.SendEmail(oMailAccount, subject, to, body, true);
                // Log start sending email
                StringBuilder sbOperation = new StringBuilder();
                sbOperation.AppendLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                sbOperation.AppendLine("Start posting to card group!");
                sbOperation.AppendLine("Subject: " + subject);
                sbOperation.AppendLine("From: " + oMailAccount.Username);
                sbOperation.AppendLine("To: " + to);
                Handler.Log(logFilePath, sbOperation.ToString());
                myNotifyIcon.BalloonTipText = "Finish posting to card group!";
            }                             
        }

        private void SendGmail(object Index)
        {
            string subject = "testing", to = "tommynguyen987@gmail.com", body = "this is testing";
            string[] postFileList = null;
            string groupsEmailPath = "";
            int index = (int)Index;
            MailAccount oMailAccount = new MailAccount();

            int delay = Handler.GetDelay(delayFilePath);
            if (delay > 0)
            {
                Thread.Sleep(delay * 60 * 60 * 1000);
            }

            string categoryPath = CategoryFolder + Categories[index] ;
            Handler.ConvertFile(categoryPath + "\\Posts");
            postFileList = Handler.GetFiles(categoryPath + "\\Posts", PATTERN_HTML);
            groupsEmailPath = categoryPath + "\\EmailList.txt";
            delayFilePath = delayFilePath.Replace(".txt", index + ".txt");
                    
            if (postFileList.Length == 0)
            {
                return;
            }

            if (index == 0)
            {
                foreach (var postFile in postFileList)
                {
                    FileInfo file = new FileInfo(postFile);
                    subject = file.Name.Replace(file.Extension, "");

                    string time = subject.Split('-')[1];
                    string temp = time.Replace(time.Substring(time.IndexOf(']')), "").Trim();
                    DateTime datetime = DateTime.Parse(temp);
                    if (datetime.Year == DateTime.Now.Year)
                    {
                        datetime = datetime.AddDays(1);
                    }

                    if (DateTime.Now > datetime)
                    {
                        file.Delete();
                        string tempWord = postFile.Replace(".html", ".docx");
                        file = new FileInfo(tempWord);
                        file.Delete();
                        string tempFolder = postFile.Replace(".html", "_files");
                        DirectoryInfo dir = new DirectoryInfo(tempFolder);
                        FileInfo[] files = dir.GetFiles();
                        foreach (var item in files)
                        {
                            item.Delete();
                        }
                        dir.Delete();
                    }
                    else
                    {
                        subject = subject.Replace("Săn Khuyến Mãi - ", "");
                        temp = subject.Substring(subject.IndexOf("]") + 1).Trim();
                        time = subject.Replace("[ ", "").Replace("] " + temp, "").Trim();
                        time = DateTime.Parse(time).ToString("dd/MM/yyyy");
                        subject = " [ " + time + " ] " + temp;
                        body = HttpUtility.HtmlDecode(File.ReadAllText(postFile, Encoding.UTF8));
                        string[] emailsList = File.ReadAllLines(groupsEmailPath)
                                                .Where(arg => !string.IsNullOrWhiteSpace(arg))
                                                .Distinct()
                                                .ToArray();

                        if (EmailMarketing.IsExistName(subject))
                        {
                            continue;
                        }

                        bool isInternetConnected = true;
                        do
                        {
                            isInternetConnected = Handler.IsAvailableNetworkActive();
                        } while (!isInternetConnected);

                        // Log start sending email
                        StringBuilder sbOperation = new StringBuilder();
                        sbOperation.AppendLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"));
                        sbOperation.AppendLine("Start posting to group!");
                        sbOperation.AppendLine("Subject: " + subject);
                        sbOperation.AppendLine("From email file: " + groupsEmailPath);
                        Handler.Log(logFilePath, sbOperation.ToString());

                        // Send email
                        Handler.SendEmail(oMailAccount, subject, emailsList, body);
                    }
                }
            }            
            myNotifyIcon.BalloonTipText = "Finish posting to group!";               
        }

        private void showMainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WindowsInSystemTray(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myNotifyIcon.Dispose();
            this.Close();
            this.Dispose();            
        }

        private void myNotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WindowsInSystemTray(false);
        }

        private void NTPEmailMarketing_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                WindowsInSystemTray(true);
            }
        }

        private void NTPEmailMarketing_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            WindowsInSystemTray(true);
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog1.FileName;
                var lines = File.ReadAllLines(txtFilePath.Text);
                foreach (var line in lines)
                {
                    int index = grvEmails.Rows.Add();
                    grvEmails.Rows[index].Cells[0].Value = (index + 1);
                    grvEmails.Rows[index].Cells[1].Value = line;
                }                
            }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            btnBrowse.Enabled = false;
            btnDelete.Enabled = false;
            backgroundWorker1.RunWorkerAsync();                        
        } 

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {            
            if(e.Cancel)
            {                
                backgroundWorker1.CancelAsync();
            }
            else
            {                
                listThreads = new Thread[7];
                for (int i = 0; i < 6; i++)
                {
                    int input = i;
                    listThreads[i] = new Thread(new ParameterizedThreadStart(SendGmail));
                    //listThreads[i].Start(0);
                }
                //listThreads[0].Start(0);

                listThreads[6] = new Thread(PostCardGroup);
                listThreads[6].Start();
            }            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            myStatus.Text = "Sending email...";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {                
                myStatus.Text = "Sending email cancelled";
            }
            else
            {
                myStatus.Text = "Done";
            }
            btnBrowse.Enabled = true;
            btnDelete.Enabled = true;
        }

        private void grvEmails_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            for (int i = 0; i < grvEmails.Rows.Count; i++)
            {
                grvEmails.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Bạn có chắc muốn xóa dòng này không?", "Xác Nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in grvEmails.Rows)
                {
                    if (row.Selected)
                    {
                        int index = row.Index;
                        grvEmails.Rows.Remove(row);
                    }
                }
                grvEmails.Refresh();
            }
        }

        private void NTPEmailMarketing_Load(object sender, EventArgs e)
        {
            myNotifyIcon.BalloonTipText = "Your application is still working" + System.Environment.NewLine + "Double click into icon to show application.";
            WindowsInSystemTray(true);            
            backgroundWorker1.RunWorkerAsync();
        }

        private void reminderPostCard_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now.Hour > 8 && DateTime.Now.Hour < 23)
            {
                PostCardGroup();
            }
        }

        private void reminderPostGroups_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now.Hour > 7 && DateTime.Now.Hour < 24)
            {
                SendGmail(0);    
            }            
        }        
    }
}
