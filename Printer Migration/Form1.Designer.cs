namespace Printer_Migration
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.bwUserImport = new System.ComponentModel.BackgroundWorker();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.tbAddPrinters = new System.Windows.Forms.TextBox();
            this.tbDeletePrinters = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pbReport = new System.Windows.Forms.ProgressBar();
            this.pbCleanUp = new System.Windows.Forms.ProgressBar();
            this.pbExecPayload = new System.Windows.Forms.ProgressBar();
            this.pbSendPayload = new System.Windows.Forms.ProgressBar();
            this.pbPingComputers = new System.Windows.Forms.ProgressBar();
            this.pbGenComputers = new System.Windows.Forms.ProgressBar();
            this.pbUserImport = new System.Windows.Forms.ProgressBar();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnRun = new System.Windows.Forms.Button();
            this.bwRun = new System.ComponentModel.BackgroundWorker();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Path to users .xls file:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel Files|*.xls*";
            this.openFileDialog1.Title = "User Text File";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(140, 63);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(234, 20);
            this.txtPath.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(380, 61);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Floor #:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Retired Printers:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(191, 39);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Printer to Add:";
            // 
            // bwUserImport
            // 
            this.bwUserImport.WorkerReportsProgress = true;
            this.bwUserImport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwUserImport_DoWork);
            this.bwUserImport.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bwUserImport_ProgressChanged);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(82, 36);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(96, 20);
            this.textBox2.TabIndex = 7;
            // 
            // tbAddPrinters
            // 
            this.tbAddPrinters.Location = new System.Drawing.Point(271, 37);
            this.tbAddPrinters.Name = "tbAddPrinters";
            this.tbAddPrinters.Size = new System.Drawing.Size(184, 20);
            this.tbAddPrinters.TabIndex = 8;
            // 
            // tbDeletePrinters
            // 
            this.tbDeletePrinters.Location = new System.Drawing.Point(140, 90);
            this.tbDeletePrinters.Name = "tbDeletePrinters";
            this.tbDeletePrinters.Size = new System.Drawing.Size(314, 20);
            this.tbDeletePrinters.TabIndex = 9;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(491, 24);
            this.menuStrip1.TabIndex = 11;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pbReport);
            this.groupBox1.Controls.Add(this.pbCleanUp);
            this.groupBox1.Controls.Add(this.pbExecPayload);
            this.groupBox1.Controls.Add(this.pbSendPayload);
            this.groupBox1.Controls.Add(this.pbPingComputers);
            this.groupBox1.Controls.Add(this.pbGenComputers);
            this.groupBox1.Controls.Add(this.pbUserImport);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Location = new System.Drawing.Point(31, 170);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(424, 250);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Details";
            // 
            // pbReport
            // 
            this.pbReport.Location = new System.Drawing.Point(160, 177);
            this.pbReport.Name = "pbReport";
            this.pbReport.Size = new System.Drawing.Size(258, 23);
            this.pbReport.TabIndex = 14;
            // 
            // pbCleanUp
            // 
            this.pbCleanUp.Location = new System.Drawing.Point(160, 206);
            this.pbCleanUp.Name = "pbCleanUp";
            this.pbCleanUp.Size = new System.Drawing.Size(258, 23);
            this.pbCleanUp.TabIndex = 13;
            // 
            // pbExecPayload
            // 
            this.pbExecPayload.Location = new System.Drawing.Point(160, 148);
            this.pbExecPayload.Name = "pbExecPayload";
            this.pbExecPayload.Size = new System.Drawing.Size(258, 23);
            this.pbExecPayload.TabIndex = 12;
            // 
            // pbSendPayload
            // 
            this.pbSendPayload.Location = new System.Drawing.Point(160, 118);
            this.pbSendPayload.Name = "pbSendPayload";
            this.pbSendPayload.Size = new System.Drawing.Size(258, 23);
            this.pbSendPayload.TabIndex = 11;
            // 
            // pbPingComputers
            // 
            this.pbPingComputers.Location = new System.Drawing.Point(160, 88);
            this.pbPingComputers.Name = "pbPingComputers";
            this.pbPingComputers.Size = new System.Drawing.Size(258, 23);
            this.pbPingComputers.TabIndex = 10;
            // 
            // pbGenComputers
            // 
            this.pbGenComputers.Location = new System.Drawing.Point(160, 58);
            this.pbGenComputers.Name = "pbGenComputers";
            this.pbGenComputers.Size = new System.Drawing.Size(258, 23);
            this.pbGenComputers.TabIndex = 9;
            // 
            // pbUserImport
            // 
            this.pbUserImport.Location = new System.Drawing.Point(160, 28);
            this.pbUserImport.Name = "pbUserImport";
            this.pbUserImport.Size = new System.Drawing.Size(258, 23);
            this.pbUserImport.TabIndex = 8;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(16, 210);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(51, 13);
            this.label10.TabIndex = 5;
            this.label10.Text = "Clean Up";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 153);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(95, 13);
            this.label9.TabIndex = 4;
            this.label9.Text = "Executing Payload";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(15, 183);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(94, 13);
            this.label12.TabIndex = 7;
            this.label12.Text = "Generating Report";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(14, 122);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(104, 13);
            this.label8.TabIndex = 3;
            this.label8.Text = "Transferring Payload";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 92);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(109, 13);
            this.label7.TabIndex = 2;
            this.label7.Text = "Accessing Computers";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 62);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(133, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Generate Computer names";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 31);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Import Users";
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(379, 128);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 13;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // bwRun
            // 
            this.bwRun.WorkerReportsProgress = true;
            this.bwRun.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwRun_DoWork);
            this.bwRun.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bwRun_ProgressChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(491, 453);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.tbDeletePrinters);
            this.Controls.Add(this.tbAddPrinters);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Printer Migration";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.ComponentModel.BackgroundWorker bwUserImport;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox tbAddPrinters;
        private System.Windows.Forms.TextBox tbDeletePrinters;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ProgressBar pbUserImport;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.ProgressBar pbReport;
        private System.Windows.Forms.ProgressBar pbCleanUp;
        private System.Windows.Forms.ProgressBar pbExecPayload;
        private System.Windows.Forms.ProgressBar pbSendPayload;
        private System.Windows.Forms.ProgressBar pbPingComputers;
        private System.Windows.Forms.ProgressBar pbGenComputers;
        private System.ComponentModel.BackgroundWorker bwRun;
    }
}

