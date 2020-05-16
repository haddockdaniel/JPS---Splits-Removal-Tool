namespace JurisUtilityBase
{
    partial class UtilityBaseMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UtilityBaseMain));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.JurisLogoImageBox = new System.Windows.Forms.PictureBox();
            this.LexisNexisLogoPictureBox = new System.Windows.Forms.PictureBox();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.listBoxCompanies = new System.Windows.Forms.ListBox();
            this.labelDescription = new System.Windows.Forms.Label();
            this.statusGroupBox = new System.Windows.Forms.GroupBox();
            this.labelCurrentStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelPercentComplete = new System.Windows.Forms.Label();
            this.OpenFileDialogOpen = new System.Windows.Forms.OpenFileDialog();
            this.btSplit = new System.Windows.Forms.Button();
            this.cbMatter = new System.Windows.Forms.ComboBox();
            this.cbClient = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).BeginInit();
            this.statusStrip.SuspendLayout();
            this.statusGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // JurisLogoImageBox
            // 
            this.JurisLogoImageBox.Image = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.Image")));
            this.JurisLogoImageBox.InitialImage = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.InitialImage")));
            this.JurisLogoImageBox.Location = new System.Drawing.Point(0, 1);
            this.JurisLogoImageBox.Name = "JurisLogoImageBox";
            this.JurisLogoImageBox.Size = new System.Drawing.Size(104, 336);
            this.JurisLogoImageBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.JurisLogoImageBox.TabIndex = 0;
            this.JurisLogoImageBox.TabStop = false;
            // 
            // LexisNexisLogoPictureBox
            // 
            this.LexisNexisLogoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LexisNexisLogoPictureBox.Image")));
            this.LexisNexisLogoPictureBox.Location = new System.Drawing.Point(8, 343);
            this.LexisNexisLogoPictureBox.Name = "LexisNexisLogoPictureBox";
            this.LexisNexisLogoPictureBox.Size = new System.Drawing.Size(96, 28);
            this.LexisNexisLogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.LexisNexisLogoPictureBox.TabIndex = 1;
            this.LexisNexisLogoPictureBox.TabStop = false;
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 379);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(1021, 22);
            this.statusStrip.TabIndex = 2;
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripStatusLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(135, 17);
            this.toolStripStatusLabel.Text = "Status: Ready to Execute";
            // 
            // listBoxCompanies
            // 
            this.listBoxCompanies.FormattingEnabled = true;
            this.listBoxCompanies.Location = new System.Drawing.Point(109, 1);
            this.listBoxCompanies.Name = "listBoxCompanies";
            this.listBoxCompanies.Size = new System.Drawing.Size(289, 43);
            this.listBoxCompanies.TabIndex = 0;
            this.listBoxCompanies.SelectedIndexChanged += new System.EventHandler(this.listBoxCompanies_SelectedIndexChanged);
            // 
            // labelDescription
            // 
            this.labelDescription.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelDescription.Location = new System.Drawing.Point(404, 1);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(274, 48);
            this.labelDescription.TabIndex = 1;
            this.labelDescription.Text = "Remove Splits from Selected Client/Matters";
            // 
            // statusGroupBox
            // 
            this.statusGroupBox.Controls.Add(this.labelCurrentStatus);
            this.statusGroupBox.Controls.Add(this.progressBar);
            this.statusGroupBox.Controls.Add(this.labelPercentComplete);
            this.statusGroupBox.Location = new System.Drawing.Point(109, 52);
            this.statusGroupBox.Name = "statusGroupBox";
            this.statusGroupBox.Size = new System.Drawing.Size(569, 73);
            this.statusGroupBox.TabIndex = 5;
            this.statusGroupBox.TabStop = false;
            this.statusGroupBox.Text = "Utility Status:";
            // 
            // labelCurrentStatus
            // 
            this.labelCurrentStatus.AutoSize = true;
            this.labelCurrentStatus.Location = new System.Drawing.Point(7, 50);
            this.labelCurrentStatus.Name = "labelCurrentStatus";
            this.labelCurrentStatus.Size = new System.Drawing.Size(77, 13);
            this.labelCurrentStatus.TabIndex = 2;
            this.labelCurrentStatus.Text = "Current Status:";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(7, 27);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(519, 20);
            this.progressBar.TabIndex = 0;
            // 
            // labelPercentComplete
            // 
            this.labelPercentComplete.AutoSize = true;
            this.labelPercentComplete.Location = new System.Drawing.Point(413, 11);
            this.labelPercentComplete.Name = "labelPercentComplete";
            this.labelPercentComplete.Size = new System.Drawing.Size(62, 13);
            this.labelPercentComplete.TabIndex = 0;
            this.labelPercentComplete.Text = "% Complete";
            // 
            // OpenFileDialogOpen
            // 
            this.OpenFileDialogOpen.FileName = "openFileDialog1";
            // 
            // btSplit
            // 
            this.btSplit.Location = new System.Drawing.Point(404, 347);
            this.btSplit.Name = "btSplit";
            this.btSplit.Size = new System.Drawing.Size(311, 24);
            this.btSplit.TabIndex = 6;
            this.btSplit.Text = "Remove Splits";
            this.btSplit.UseVisualStyleBackColor = true;
            this.btSplit.Click += new System.EventHandler(this.button1_Click);
            // 
            // cbMatter
            // 
            this.cbMatter.FormattingEnabled = true;
            this.cbMatter.Location = new System.Drawing.Point(272, 154);
            this.cbMatter.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cbMatter.Name = "cbMatter";
            this.cbMatter.Size = new System.Drawing.Size(406, 21);
            this.cbMatter.TabIndex = 8;
            this.cbMatter.SelectedIndexChanged += new System.EventHandler(this.cbMatter_SelectedIndexChanged);
            // 
            // cbClient
            // 
            this.cbClient.FormattingEnabled = true;
            this.cbClient.Location = new System.Drawing.Point(272, 130);
            this.cbClient.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cbClient.Name = "cbClient";
            this.cbClient.Size = new System.Drawing.Size(406, 21);
            this.cbClient.TabIndex = 9;
            this.cbClient.SelectedIndexChanged += new System.EventHandler(this.cbClient_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(191, 133);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Select Client(s)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(187, 157);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Select Matter(s)";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Location = new System.Drawing.Point(109, 183);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(901, 154);
            this.dataGridView1.TabIndex = 13;
            // 
            // UtilityBaseMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1021, 401);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbClient);
            this.Controls.Add(this.cbMatter);
            this.Controls.Add(this.btSplit);
            this.Controls.Add(this.statusGroupBox);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.listBoxCompanies);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.LexisNexisLogoPictureBox);
            this.Controls.Add(this.JurisLogoImageBox);
            this.ForeColor = System.Drawing.SystemColors.WindowText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UtilityBaseMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "JPS - Split Removal Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).EndInit();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.statusGroupBox.ResumeLayout(false);
            this.statusGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox JurisLogoImageBox;
        private System.Windows.Forms.PictureBox LexisNexisLogoPictureBox;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ListBox listBoxCompanies;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.GroupBox statusGroupBox;
        private System.Windows.Forms.Label labelCurrentStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label labelPercentComplete;
        private System.Windows.Forms.OpenFileDialog OpenFileDialogOpen;
        private System.Windows.Forms.Button btSplit;
        private System.Windows.Forms.ComboBox cbMatter;
        private System.Windows.Forms.ComboBox cbClient;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

