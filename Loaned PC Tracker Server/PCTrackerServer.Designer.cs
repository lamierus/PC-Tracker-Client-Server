﻿namespace Loaned_PC_Tracker_Server {
    public partial class PCTrackerServerForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.bgwLoadSites = new System.ComponentModel.BackgroundWorker();
            this.bgwLoadPCs = new System.ComponentModel.BackgroundWorker();
            this.bgwAutoSave = new System.ComponentModel.BackgroundWorker();
            this.tbLog = new System.Windows.Forms.TextBox();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoSaveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.quitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.otherToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testBroadcastToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // bgwLoadSites
            // 
            this.bgwLoadSites.WorkerReportsProgress = true;
            this.bgwLoadSites.WorkerSupportsCancellation = true;
            this.bgwLoadSites.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwLoadSites_DoWork);
            this.bgwLoadSites.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgwLoadSites_ProgressChanged);
            this.bgwLoadSites.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgwLoadSites_RunWorkerCompleted);
            // 
            // bgwLoadPCs
            // 
            this.bgwLoadPCs.WorkerReportsProgress = true;
            this.bgwLoadPCs.WorkerSupportsCancellation = true;
            this.bgwLoadPCs.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwLoadPCs_DoWork);
            this.bgwLoadPCs.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bgwLoadPCs_ProgressChanged);
            this.bgwLoadPCs.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgwLoadPCs_RunWorkerCompleted);
            // 
            // bgwAutoSave
            // 
            this.bgwAutoSave.WorkerReportsProgress = true;
            this.bgwAutoSave.WorkerSupportsCancellation = true;
            this.bgwAutoSave.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwAutoSave_DoWork);
            this.bgwAutoSave.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bgwAutoSave_RunWorkerCompleted);
            // 
            // tbLog
            // 
            this.tbLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbLog.Location = new System.Drawing.Point(16, 33);
            this.tbLog.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbLog.Multiline = true;
            this.tbLog.Name = "tbLog";
            this.tbLog.ReadOnly = true;
            this.tbLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tbLog.Size = new System.Drawing.Size(727, 259);
            this.tbLog.TabIndex = 0;
            // 
            // menuStrip
            // 
            this.menuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.otherToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Padding = new System.Windows.Forms.Padding(8, 2, 0, 2);
            this.menuStrip.Size = new System.Drawing.Size(760, 28);
            this.menuStrip.TabIndex = 1;
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.autoSaveToolStripMenuItem,
            this.quitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(151, 26);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // autoSaveToolStripMenuItem
            // 
            this.autoSaveToolStripMenuItem.Checked = true;
            this.autoSaveToolStripMenuItem.CheckOnClick = true;
            this.autoSaveToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.autoSaveToolStripMenuItem.Name = "autoSaveToolStripMenuItem";
            this.autoSaveToolStripMenuItem.Size = new System.Drawing.Size(151, 26);
            this.autoSaveToolStripMenuItem.Text = "Auto Save";
            this.autoSaveToolStripMenuItem.Click += new System.EventHandler(this.autoSaveToolStripMenuItem_Click);
            // 
            // quitToolStripMenuItem
            // 
            this.quitToolStripMenuItem.Name = "quitToolStripMenuItem";
            this.quitToolStripMenuItem.Size = new System.Drawing.Size(151, 26);
            this.quitToolStripMenuItem.Text = "Quit";
            this.quitToolStripMenuItem.Click += new System.EventHandler(this.quitToolStripMenuItem_Click);
            // 
            // otherToolStripMenuItem
            // 
            this.otherToolStripMenuItem.Checked = true;
            this.otherToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.otherToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testBroadcastToolStripMenuItem});
            this.otherToolStripMenuItem.Name = "otherToolStripMenuItem";
            this.otherToolStripMenuItem.Size = new System.Drawing.Size(58, 24);
            this.otherToolStripMenuItem.Text = "Other";
            // 
            // testBroadcastToolStripMenuItem
            // 
            this.testBroadcastToolStripMenuItem.Name = "testBroadcastToolStripMenuItem";
            this.testBroadcastToolStripMenuItem.Size = new System.Drawing.Size(180, 26);
            this.testBroadcastToolStripMenuItem.Text = "Test Broadcast";
            this.testBroadcastToolStripMenuItem.Click += new System.EventHandler(this.testBroadcastToolStripMenuItem_Click);
            // 
            // PCTrackerServerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(760, 308);
            this.Controls.Add(this.tbLog);
            this.Controls.Add(this.menuStrip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MainMenuStrip = this.menuStrip;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "PCTrackerServerForm";
            this.Text = "6";
            this.Activated += new System.EventHandler(this.PCTrackerServerForm_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.PCTrackerServerForm_Closing);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.ComponentModel.BackgroundWorker bgwLoadSites;
        private System.ComponentModel.BackgroundWorker bgwLoadPCs;
        private System.ComponentModel.BackgroundWorker bgwAutoSave;
        private System.Windows.Forms.TextBox tbLog;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem quitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem otherToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem testBroadcastToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoSaveToolStripMenuItem;
    }
}

