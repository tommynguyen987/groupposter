namespace NTPEmailMarketing
{
    partial class GroupPoster
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GroupPoster));
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.grvEmails = new System.Windows.Forms.DataGridView();
            this.ColSTT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColEmail = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.myStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.myProgress = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.btnDelete = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.myNotifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.myContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.showMainToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reminderPostCard = new System.Windows.Forms.Timer(this.components);
            this.reminderPostGroups = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.grvEmails)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.toolStripContainer1.BottomToolStripPanel.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            this.myContextMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtFilePath
            // 
            this.txtFilePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtFilePath.Location = new System.Drawing.Point(132, 16);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(283, 20);
            this.txtFilePath.TabIndex = 2;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(432, 15);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 3;
            this.btnBrowse.Text = "Duyệt File";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnSend
            // 
            this.btnSend.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSend.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSend.Location = new System.Drawing.Point(432, 95);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 45);
            this.btnSend.TabIndex = 3;
            this.btnSend.Text = "SEND";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Nhập danh sách email";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // grvEmails
            // 
            this.grvEmails.AllowUserToAddRows = false;
            this.grvEmails.AllowUserToDeleteRows = false;
            this.grvEmails.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grvEmails.BackgroundColor = System.Drawing.Color.GhostWhite;
            this.grvEmails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.grvEmails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grvEmails.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColSTT,
            this.ColEmail,
            this.ColStatus});
            this.grvEmails.GridColor = System.Drawing.SystemColors.Control;
            this.grvEmails.Location = new System.Drawing.Point(16, 55);
            this.grvEmails.MultiSelect = false;
            this.grvEmails.Name = "grvEmails";
            this.grvEmails.ReadOnly = true;
            this.grvEmails.RowHeadersVisible = false;
            this.grvEmails.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.grvEmails.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grvEmails.ShowCellErrors = false;
            this.grvEmails.ShowEditingIcon = false;
            this.grvEmails.ShowRowErrors = false;
            this.grvEmails.Size = new System.Drawing.Size(399, 283);
            this.grvEmails.TabIndex = 5;
            this.grvEmails.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.grvEmails_RowsRemoved);
            // 
            // ColSTT
            // 
            this.ColSTT.FillWeight = 106.599F;
            this.ColSTT.HeaderText = "STT";
            this.ColSTT.MinimumWidth = 20;
            this.ColSTT.Name = "ColSTT";
            this.ColSTT.ReadOnly = true;
            this.ColSTT.Width = 80;
            // 
            // ColEmail
            // 
            this.ColEmail.FillWeight = 96.70051F;
            this.ColEmail.HeaderText = "Email";
            this.ColEmail.MinimumWidth = 100;
            this.ColEmail.Name = "ColEmail";
            this.ColEmail.ReadOnly = true;
            this.ColEmail.Width = 200;
            // 
            // ColStatus
            // 
            this.ColStatus.FillWeight = 96.70051F;
            this.ColStatus.HeaderText = "Status";
            this.ColStatus.MinimumWidth = 50;
            this.ColStatus.Name = "ColStatus";
            this.ColStatus.ReadOnly = true;
            this.ColStatus.Width = 120;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.myStatus,
            this.myProgress});
            this.statusStrip1.Location = new System.Drawing.Point(0, 0);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.statusStrip1.Size = new System.Drawing.Size(523, 26);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // myStatus
            // 
            this.myStatus.Name = "myStatus";
            this.myStatus.Size = new System.Drawing.Size(39, 21);
            this.myStatus.Text = "Ready";
            // 
            // myProgress
            // 
            this.myProgress.Margin = new System.Windows.Forms.Padding(5, 3, 1, 3);
            this.myProgress.Name = "myProgress";
            this.myProgress.Size = new System.Drawing.Size(300, 20);
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.BottomToolStripPanel
            // 
            this.toolStripContainer1.BottomToolStripPanel.Controls.Add(this.statusStrip1);
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.AutoScroll = true;
            this.toolStripContainer1.ContentPanel.Controls.Add(this.grvEmails);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnSend);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnDelete);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnBrowse);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.txtFilePath);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(523, 356);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.LeftToolStripPanelVisible = false;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.RightToolStripPanelVisible = false;
            this.toolStripContainer1.Size = new System.Drawing.Size(523, 382);
            this.toolStripContainer1.TabIndex = 7;
            this.toolStripContainer1.Text = "toolStripContainer1";
            this.toolStripContainer1.TopToolStripPanelVisible = false;
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDelete.Location = new System.Drawing.Point(432, 54);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 3;
            this.btnDelete.Text = "Xóa Email";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // myNotifyIcon
            // 
            this.myNotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.myNotifyIcon.BalloonTipTitle = "Group Poster";
            this.myNotifyIcon.ContextMenuStrip = this.myContextMenuStrip;
            this.myNotifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("myNotifyIcon.Icon")));
            this.myNotifyIcon.Text = "Double click into icon to show application";
            this.myNotifyIcon.Visible = true;
            this.myNotifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.myNotifyIcon_MouseDoubleClick);
            // 
            // myContextMenuStrip
            // 
            this.myContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.showMainToolStripMenuItem,
            this.toolStripSeparator1,
            this.exitToolStripMenuItem});
            this.myContextMenuStrip.Name = "myContextMenuStrip";
            this.myContextMenuStrip.Size = new System.Drawing.Size(184, 54);
            // 
            // showMainToolStripMenuItem
            // 
            this.showMainToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.showMainToolStripMenuItem.Name = "showMainToolStripMenuItem";
            this.showMainToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.showMainToolStripMenuItem.Text = "Open Main Window";
            this.showMainToolStripMenuItem.Click += new System.EventHandler(this.showMainToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(180, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // reminderPostCard
            // 
            this.reminderPostCard.Enabled = true;
            this.reminderPostCard.Interval = 600000;
            this.reminderPostCard.Tick += new System.EventHandler(this.reminderPostCard_Tick);
            // 
            // reminderPostGroups
            // 
            this.reminderPostGroups.Enabled = true;
            this.reminderPostGroups.Interval = 14400000;
            this.reminderPostGroups.Tick += new System.EventHandler(this.reminderPostGroups_Tick);
            // 
            // GroupPoster
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(523, 382);
            this.Controls.Add(this.toolStripContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "GroupPoster";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "NTP Email Marketing";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NTPEmailMarketing_FormClosing);
            this.Load += new System.EventHandler(this.NTPEmailMarketing_Load);
            this.Resize += new System.EventHandler(this.NTPEmailMarketing_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.grvEmails)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolStripContainer1.BottomToolStripPanel.ResumeLayout(false);
            this.toolStripContainer1.BottomToolStripPanel.PerformLayout();
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.ContentPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            this.myContextMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView grvEmails;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar myProgress;
        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSTT;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColEmail;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColStatus;
        private System.Windows.Forms.ToolStripStatusLabel myStatus;
        private System.Windows.Forms.NotifyIcon myNotifyIcon;
        private System.Windows.Forms.ContextMenuStrip myContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem showMainToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.Timer reminderPostCard;
        private System.Windows.Forms.Timer reminderPostGroups;
    }
}

