namespace HealthCalendarAdmin
{
    partial class HealthCalendarAdmin
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HealthCalendarAdmin));
            this.txtboxSearchFirstname = new System.Windows.Forms.TextBox();
            this.lblSearchFirstname = new System.Windows.Forms.Label();
            this.dgvSubscribers = new System.Windows.Forms.DataGridView();
            this.btnGoogleEmailClear = new System.Windows.Forms.Button();
            this.btnGoogleEmailUpdate = new System.Windows.Forms.Button();
            this.txtboxGoogleEmail = new System.Windows.Forms.TextBox();
            this.lblGoogleEmail = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pnlFooter = new System.Windows.Forms.Panel();
            this.lblPercent = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnGoogleCreateShare = new System.Windows.Forms.Button();
            this.btnSampleGoogleData = new System.Windows.Forms.Button();
            this.btnClearGoogleCalendar = new System.Windows.Forms.Button();
            this.bgWorker = new System.ComponentModel.BackgroundWorker();
            this.btnDeleteGoogleCalendar = new System.Windows.Forms.Button();
            this.txtboxSearchLastname = new System.Windows.Forms.TextBox();
            this.groupBoxSearch = new System.Windows.Forms.GroupBox();
            this.lblSearchMainIdentifier = new System.Windows.Forms.Label();
            this.txtboxSearchMainIdentifier = new System.Windows.Forms.TextBox();
            this.lblSearchOccupation = new System.Windows.Forms.Label();
            this.comboBoxOccupation = new System.Windows.Forms.ComboBox();
            this.lblSearchTitle = new System.Windows.Forms.Label();
            this.comboBoxTitle = new System.Windows.Forms.ComboBox();
            this.lblSearchSex = new System.Windows.Forms.Label();
            this.lblSearchLastname = new System.Windows.Forms.Label();
            this.comboBoxSex = new System.Windows.Forms.ComboBox();
            this.btnClearSearch = new System.Windows.Forms.Button();
            this.groupBoxGoogle = new System.Windows.Forms.GroupBox();
            this.groupBoxNHSNet = new System.Windows.Forms.GroupBox();
            this.lblNHSNetEmail = new System.Windows.Forms.Label();
            this.txtboxNHSNetEmail = new System.Windows.Forms.TextBox();
            this.btnDeleteNHSNetCalendar = new System.Windows.Forms.Button();
            this.btnNHSNetEmailClear = new System.Windows.Forms.Button();
            this.btnNHSNetEmailUpdate = new System.Windows.Forms.Button();
            this.btnClearNHSNetCalendar = new System.Windows.Forms.Button();
            this.btnSampleNHSNetData = new System.Windows.Forms.Button();
            this.btnNHSNetCreateShare = new System.Windows.Forms.Button();
            this.groupBoxExchange = new System.Windows.Forms.GroupBox();
            this.lblExchangeEmail = new System.Windows.Forms.Label();
            this.txtboxExchangeEmail = new System.Windows.Forms.TextBox();
            this.btnDeleteExchangeCalendar = new System.Windows.Forms.Button();
            this.btnExchangeEmailClear = new System.Windows.Forms.Button();
            this.btnExchangeEmailUpdate = new System.Windows.Forms.Button();
            this.btnClearExchangeCalendar = new System.Windows.Forms.Button();
            this.btnSampleExchangeData = new System.Windows.Forms.Button();
            this.btnExchangeCreateShare = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubscribers)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.pnlFooter.SuspendLayout();
            this.groupBoxSearch.SuspendLayout();
            this.groupBoxGoogle.SuspendLayout();
            this.groupBoxNHSNet.SuspendLayout();
            this.groupBoxExchange.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtboxSearchFirstname
            // 
            this.txtboxSearchFirstname.AccessibleDescription = "First name search box";
            this.txtboxSearchFirstname.AccessibleName = "First name search box";
            this.txtboxSearchFirstname.AccessibleRole = System.Windows.Forms.AccessibleRole.Text;
            this.txtboxSearchFirstname.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxSearchFirstname.Location = new System.Drawing.Point(116, 11);
            this.txtboxSearchFirstname.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxSearchFirstname.Name = "txtboxSearchFirstname";
            this.txtboxSearchFirstname.ShortcutsEnabled = false;
            this.txtboxSearchFirstname.Size = new System.Drawing.Size(130, 19);
            this.txtboxSearchFirstname.TabIndex = 1;
            this.txtboxSearchFirstname.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // lblSearchFirstname
            // 
            this.lblSearchFirstname.AutoSize = true;
            this.lblSearchFirstname.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchFirstname.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchFirstname.Location = new System.Drawing.Point(54, 11);
            this.lblSearchFirstname.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchFirstname.Name = "lblSearchFirstname";
            this.lblSearchFirstname.Size = new System.Drawing.Size(57, 13);
            this.lblSearchFirstname.TabIndex = 40;
            this.lblSearchFirstname.Text = "First Name";
            // 
            // dgvSubscribers
            // 
            this.dgvSubscribers.AccessibleDescription = "Filtered list of Subscribers";
            this.dgvSubscribers.AccessibleName = "Calendar Subscriber Table";
            this.dgvSubscribers.AccessibleRole = System.Windows.Forms.AccessibleRole.Table;
            this.dgvSubscribers.AllowUserToAddRows = false;
            this.dgvSubscribers.AllowUserToDeleteRows = false;
            this.dgvSubscribers.AllowUserToResizeRows = false;
            this.dgvSubscribers.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgvSubscribers.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Sunken;
            this.dgvSubscribers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.Padding = new System.Windows.Forms.Padding(0, 4, 0, 4);
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvSubscribers.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvSubscribers.Location = new System.Drawing.Point(9, 74);
            this.dgvSubscribers.Margin = new System.Windows.Forms.Padding(1);
            this.dgvSubscribers.MultiSelect = false;
            this.dgvSubscribers.Name = "dgvSubscribers";
            this.dgvSubscribers.ReadOnly = true;
            this.dgvSubscribers.RowTemplate.Height = 25;
            this.dgvSubscribers.ShowCellToolTips = false;
            this.dgvSubscribers.Size = new System.Drawing.Size(742, 258);
            this.dgvSubscribers.TabIndex = 8;
            this.dgvSubscribers.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSubscribers_CellContentClick);
            this.dgvSubscribers.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSubscribers_RowEnter);
            this.dgvSubscribers.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvSubscribers_RowHeaderMouseClick);
            // 
            // btnGoogleEmailClear
            // 
            this.btnGoogleEmailClear.BackColor = System.Drawing.Color.LightGray;
            this.btnGoogleEmailClear.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnGoogleEmailClear.BackgroundImage")));
            this.btnGoogleEmailClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnGoogleEmailClear.Enabled = false;
            this.btnGoogleEmailClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGoogleEmailClear.ForeColor = System.Drawing.Color.White;
            this.btnGoogleEmailClear.Location = new System.Drawing.Point(36, 15);
            this.btnGoogleEmailClear.Margin = new System.Windows.Forms.Padding(1);
            this.btnGoogleEmailClear.Name = "btnGoogleEmailClear";
            this.btnGoogleEmailClear.Size = new System.Drawing.Size(29, 27);
            this.btnGoogleEmailClear.TabIndex = 25;
            this.btnGoogleEmailClear.UseVisualStyleBackColor = false;
            this.btnGoogleEmailClear.Visible = false;
            this.btnGoogleEmailClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnGoogleEmailUpdate
            // 
            this.btnGoogleEmailUpdate.BackColor = System.Drawing.Color.LightGray;
            this.btnGoogleEmailUpdate.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnGoogleEmailUpdate.BackgroundImage")));
            this.btnGoogleEmailUpdate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnGoogleEmailUpdate.Enabled = false;
            this.btnGoogleEmailUpdate.FlatAppearance.BorderSize = 0;
            this.btnGoogleEmailUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGoogleEmailUpdate.ForeColor = System.Drawing.Color.White;
            this.btnGoogleEmailUpdate.Location = new System.Drawing.Point(6, 15);
            this.btnGoogleEmailUpdate.Margin = new System.Windows.Forms.Padding(1);
            this.btnGoogleEmailUpdate.Name = "btnGoogleEmailUpdate";
            this.btnGoogleEmailUpdate.Size = new System.Drawing.Size(29, 27);
            this.btnGoogleEmailUpdate.TabIndex = 24;
            this.btnGoogleEmailUpdate.UseVisualStyleBackColor = false;
            this.btnGoogleEmailUpdate.Visible = false;
            this.btnGoogleEmailUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // txtboxGoogleEmail
            // 
            this.txtboxGoogleEmail.Enabled = false;
            this.txtboxGoogleEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxGoogleEmail.Location = new System.Drawing.Point(49, 46);
            this.txtboxGoogleEmail.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxGoogleEmail.Name = "txtboxGoogleEmail";
            this.txtboxGoogleEmail.Size = new System.Drawing.Size(144, 20);
            this.txtboxGoogleEmail.TabIndex = 23;
            this.txtboxGoogleEmail.Visible = false;
            this.txtboxGoogleEmail.TextChanged += new System.EventHandler(this.txtboxGoogleEmail_TextChanged);
            // 
            // lblGoogleEmail
            // 
            this.lblGoogleEmail.AutoSize = true;
            this.lblGoogleEmail.BackColor = System.Drawing.Color.Transparent;
            this.lblGoogleEmail.Enabled = false;
            this.lblGoogleEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblGoogleEmail.Location = new System.Drawing.Point(6, 46);
            this.lblGoogleEmail.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblGoogleEmail.Name = "lblGoogleEmail";
            this.lblGoogleEmail.Size = new System.Drawing.Size(33, 13);
            this.lblGoogleEmail.TabIndex = 0;
            this.lblGoogleEmail.Text = "Gmail";
            this.lblGoogleEmail.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(746, 354);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 42;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // pnlFooter
            // 
            this.pnlFooter.BackColor = System.Drawing.Color.LightGray;
            this.pnlFooter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlFooter.Controls.Add(this.lblPercent);
            this.pnlFooter.Controls.Add(this.progressBar);
            this.pnlFooter.Controls.Add(this.textBox2);
            this.pnlFooter.Controls.Add(this.textBox1);
            this.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlFooter.Location = new System.Drawing.Point(0, 454);
            this.pnlFooter.Margin = new System.Windows.Forms.Padding(2);
            this.pnlFooter.Name = "pnlFooter";
            this.pnlFooter.Size = new System.Drawing.Size(789, 24);
            this.pnlFooter.TabIndex = 0;
            // 
            // lblPercent
            // 
            this.lblPercent.AutoSize = true;
            this.lblPercent.Location = new System.Drawing.Point(560, 5);
            this.lblPercent.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblPercent.Name = "lblPercent";
            this.lblPercent.Size = new System.Drawing.Size(0, 13);
            this.lblPercent.TabIndex = 0;
            // 
            // progressBar
            // 
            this.progressBar.ForeColor = System.Drawing.Color.OrangeRed;
            this.progressBar.Location = new System.Drawing.Point(186, 0);
            this.progressBar.Margin = new System.Windows.Forms.Padding(2);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(371, 22);
            this.progressBar.TabIndex = 0;
            this.progressBar.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.LightGray;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Dock = System.Windows.Forms.DockStyle.Right;
            this.textBox2.Font = new System.Drawing.Font("Segoe UI", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.ForeColor = System.Drawing.Color.Black;
            this.textBox2.Location = new System.Drawing.Point(667, 0);
            this.textBox2.Margin = new System.Windows.Forms.Padding(2);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(120, 18);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "Loch Roag Limited";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.LightGray;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Left;
            this.textBox1.Font = new System.Drawing.Font("Segoe UI", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.ForeColor = System.Drawing.Color.Black;
            this.textBox1.Location = new System.Drawing.Point(0, 0);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(109, 18);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "Health Calendar";
            // 
            // btnGoogleCreateShare
            // 
            this.btnGoogleCreateShare.BackColor = System.Drawing.Color.LightGray;
            this.btnGoogleCreateShare.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnGoogleCreateShare.BackgroundImage")));
            this.btnGoogleCreateShare.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnGoogleCreateShare.Enabled = false;
            this.btnGoogleCreateShare.FlatAppearance.BorderSize = 0;
            this.btnGoogleCreateShare.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGoogleCreateShare.ForeColor = System.Drawing.Color.White;
            this.btnGoogleCreateShare.Location = new System.Drawing.Point(68, 17);
            this.btnGoogleCreateShare.Margin = new System.Windows.Forms.Padding(1);
            this.btnGoogleCreateShare.Name = "btnGoogleCreateShare";
            this.btnGoogleCreateShare.Size = new System.Drawing.Size(29, 27);
            this.btnGoogleCreateShare.TabIndex = 26;
            this.btnGoogleCreateShare.UseVisualStyleBackColor = false;
            this.btnGoogleCreateShare.Visible = false;
            this.btnGoogleCreateShare.Click += new System.EventHandler(this.btnGoogleCreateShare_Click);
            // 
            // btnSampleGoogleData
            // 
            this.btnSampleGoogleData.BackColor = System.Drawing.Color.LightGray;
            this.btnSampleGoogleData.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSampleGoogleData.BackgroundImage")));
            this.btnSampleGoogleData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSampleGoogleData.Enabled = false;
            this.btnSampleGoogleData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSampleGoogleData.ForeColor = System.Drawing.Color.White;
            this.btnSampleGoogleData.Location = new System.Drawing.Point(99, 17);
            this.btnSampleGoogleData.Margin = new System.Windows.Forms.Padding(1);
            this.btnSampleGoogleData.Name = "btnSampleGoogleData";
            this.btnSampleGoogleData.Size = new System.Drawing.Size(29, 27);
            this.btnSampleGoogleData.TabIndex = 27;
            this.btnSampleGoogleData.UseVisualStyleBackColor = false;
            this.btnSampleGoogleData.Visible = false;
            this.btnSampleGoogleData.Click += new System.EventHandler(this.btnSampleGoogleData_Click);
            // 
            // btnClearGoogleCalendar
            // 
            this.btnClearGoogleCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnClearGoogleCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClearGoogleCalendar.BackgroundImage")));
            this.btnClearGoogleCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClearGoogleCalendar.Enabled = false;
            this.btnClearGoogleCalendar.FlatAppearance.BorderSize = 0;
            this.btnClearGoogleCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearGoogleCalendar.ForeColor = System.Drawing.Color.White;
            this.btnClearGoogleCalendar.Location = new System.Drawing.Point(130, 16);
            this.btnClearGoogleCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnClearGoogleCalendar.Name = "btnClearGoogleCalendar";
            this.btnClearGoogleCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnClearGoogleCalendar.TabIndex = 28;
            this.btnClearGoogleCalendar.UseVisualStyleBackColor = false;
            this.btnClearGoogleCalendar.Visible = false;
            this.btnClearGoogleCalendar.Click += new System.EventHandler(this.btnClearGoogleCalendar_Click);
            // 
            // btnDeleteGoogleCalendar
            // 
            this.btnDeleteGoogleCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnDeleteGoogleCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDeleteGoogleCalendar.BackgroundImage")));
            this.btnDeleteGoogleCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDeleteGoogleCalendar.Enabled = false;
            this.btnDeleteGoogleCalendar.FlatAppearance.BorderSize = 0;
            this.btnDeleteGoogleCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteGoogleCalendar.ForeColor = System.Drawing.Color.White;
            this.btnDeleteGoogleCalendar.Location = new System.Drawing.Point(162, 17);
            this.btnDeleteGoogleCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnDeleteGoogleCalendar.Name = "btnDeleteGoogleCalendar";
            this.btnDeleteGoogleCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnDeleteGoogleCalendar.TabIndex = 29;
            this.btnDeleteGoogleCalendar.UseVisualStyleBackColor = false;
            this.btnDeleteGoogleCalendar.Visible = false;
            this.btnDeleteGoogleCalendar.Click += new System.EventHandler(this.btnDeleteGoogleCalendar_Click);
            // 
            // txtboxSearchLastname
            // 
            this.txtboxSearchLastname.AccessibleDescription = "Last name search box";
            this.txtboxSearchLastname.AccessibleName = "Last name search box";
            this.txtboxSearchLastname.AccessibleRole = System.Windows.Forms.AccessibleRole.Text;
            this.txtboxSearchLastname.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxSearchLastname.Location = new System.Drawing.Point(116, 38);
            this.txtboxSearchLastname.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxSearchLastname.Name = "txtboxSearchLastname";
            this.txtboxSearchLastname.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtboxSearchLastname.ShortcutsEnabled = false;
            this.txtboxSearchLastname.Size = new System.Drawing.Size(130, 19);
            this.txtboxSearchLastname.TabIndex = 2;
            this.txtboxSearchLastname.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // groupBoxSearch
            // 
            this.groupBoxSearch.AccessibleDescription = "SearchGroup";
            this.groupBoxSearch.AccessibleName = "Search";
            this.groupBoxSearch.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
            this.groupBoxSearch.Controls.Add(this.lblSearchMainIdentifier);
            this.groupBoxSearch.Controls.Add(this.txtboxSearchMainIdentifier);
            this.groupBoxSearch.Controls.Add(this.lblSearchOccupation);
            this.groupBoxSearch.Controls.Add(this.comboBoxOccupation);
            this.groupBoxSearch.Controls.Add(this.lblSearchTitle);
            this.groupBoxSearch.Controls.Add(this.comboBoxTitle);
            this.groupBoxSearch.Controls.Add(this.lblSearchSex);
            this.groupBoxSearch.Controls.Add(this.lblSearchLastname);
            this.groupBoxSearch.Controls.Add(this.comboBoxSex);
            this.groupBoxSearch.Controls.Add(this.btnClearSearch);
            this.groupBoxSearch.Controls.Add(this.txtboxSearchFirstname);
            this.groupBoxSearch.Controls.Add(this.txtboxSearchLastname);
            this.groupBoxSearch.Controls.Add(this.lblSearchFirstname);
            this.groupBoxSearch.Location = new System.Drawing.Point(9, 4);
            this.groupBoxSearch.Name = "groupBoxSearch";
            this.groupBoxSearch.Size = new System.Drawing.Size(741, 64);
            this.groupBoxSearch.TabIndex = 44;
            this.groupBoxSearch.TabStop = false;
            this.groupBoxSearch.Text = "Search";
            // 
            // lblSearchMainIdentifier
            // 
            this.lblSearchMainIdentifier.AutoSize = true;
            this.lblSearchMainIdentifier.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchMainIdentifier.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchMainIdentifier.Location = new System.Drawing.Point(384, 44);
            this.lblSearchMainIdentifier.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchMainIdentifier.Name = "lblSearchMainIdentifier";
            this.lblSearchMainIdentifier.Size = new System.Drawing.Size(73, 13);
            this.lblSearchMainIdentifier.TabIndex = 54;
            this.lblSearchMainIdentifier.Text = "Main Identifier";
            // 
            // txtboxSearchMainIdentifier
            // 
            this.txtboxSearchMainIdentifier.AccessibleDescription = "Main Identifier Search Box";
            this.txtboxSearchMainIdentifier.AccessibleName = "Main Identifier Search Box";
            this.txtboxSearchMainIdentifier.AccessibleRole = System.Windows.Forms.AccessibleRole.Text;
            this.txtboxSearchMainIdentifier.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxSearchMainIdentifier.Location = new System.Drawing.Point(474, 40);
            this.txtboxSearchMainIdentifier.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxSearchMainIdentifier.Name = "txtboxSearchMainIdentifier";
            this.txtboxSearchMainIdentifier.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtboxSearchMainIdentifier.ShortcutsEnabled = false;
            this.txtboxSearchMainIdentifier.Size = new System.Drawing.Size(158, 19);
            this.txtboxSearchMainIdentifier.TabIndex = 6;
            this.txtboxSearchMainIdentifier.TextChanged += new System.EventHandler(this.txtboxSearchMainIdentifier_TextChanged);
            // 
            // lblSearchOccupation
            // 
            this.lblSearchOccupation.AutoSize = true;
            this.lblSearchOccupation.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchOccupation.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchOccupation.Location = new System.Drawing.Point(384, 10);
            this.lblSearchOccupation.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchOccupation.Name = "lblSearchOccupation";
            this.lblSearchOccupation.Size = new System.Drawing.Size(62, 13);
            this.lblSearchOccupation.TabIndex = 52;
            this.lblSearchOccupation.Text = "Occupation";
            // 
            // comboBoxOccupation
            // 
            this.comboBoxOccupation.AccessibleDescription = "Occupation Combo Box Search";
            this.comboBoxOccupation.AccessibleName = "Occupation Combo Box Search";
            this.comboBoxOccupation.AccessibleRole = System.Windows.Forms.AccessibleRole.ComboBox;
            this.comboBoxOccupation.FormattingEnabled = true;
            this.comboBoxOccupation.Items.AddRange(new object[] {
            "",
            "Anaesthetist",
            "Assistant Psychologist",
            "Audiologist",
            "Bio-medical",
            "Clerical Worker",
            "Clinical Assistant",
            "Clinical Nurse Specialist",
            "Clinical Psychologist",
            "Community Nurse",
            "Community Paediatrician",
            "Consultant",
            "Dentist",
            "Dietician",
            "Doctor",
            "General dental practitioner",
            "General Medical Practitioner",
            "Healthcare Assistant",
            "Health Care Support Worker",
            "Healthcare Practitioner",
            "Health Care Professional",
            "Health Visitor",
            "Hospital administrator",
            "Hospital Manager",
            "Hospital Consultant",
            "Hospital Nurse",
            "Hospital Registrar",
            "Immunodeficiency Consultant",
            "Immunodeficiency Nurse",
            "IT professional",
            "Medical Secretary (Clerical and Administrative)",
            "Medical Staff",
            "Midwife",
            "Macmillan Nurse",
            "Not Specified",
            "Not Specified",
            "Nurse Practitioner",
            "Nurse Manager",
            "Occupational Therapist",
            "Orthoptist",
            "Orthotist",
            "Other",
            "Other therapist",
            "Pharmacist",
            "Physiotherapist",
            "Podiatrist",
            "Practice Health Care Professional",
            "Radiographer",
            "Radiologist",
            "Receptionist",
            "School Nurse",
            "SHO/Registrar",
            "Sexual Health Fam Plan DR",
            "Social Worker",
            "Solicitor",
            "Specialist Nursing - Palliative/Respite Care",
            "Specialist Registrar",
            "Speech and Language Therapist",
            "Staff Grade",
            "Staff Nurse",
            "Support Worker",
            "Technical Instructor",
            "Operating Department Practitioner",
            "Specialist Practitioner"});
            this.comboBoxOccupation.Location = new System.Drawing.Point(474, 11);
            this.comboBoxOccupation.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxOccupation.Name = "comboBoxOccupation";
            this.comboBoxOccupation.Size = new System.Drawing.Size(158, 21);
            this.comboBoxOccupation.TabIndex = 5;
            this.comboBoxOccupation.SelectedIndexChanged += new System.EventHandler(this.comboBoxOccupation_SelectedIndexChanged);
            // 
            // lblSearchTitle
            // 
            this.lblSearchTitle.AutoSize = true;
            this.lblSearchTitle.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchTitle.Location = new System.Drawing.Point(260, 40);
            this.lblSearchTitle.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchTitle.Name = "lblSearchTitle";
            this.lblSearchTitle.Size = new System.Drawing.Size(27, 13);
            this.lblSearchTitle.TabIndex = 50;
            this.lblSearchTitle.Text = "Title";
            // 
            // comboBoxTitle
            // 
            this.comboBoxTitle.AccessibleDescription = "Title Combo Box Search";
            this.comboBoxTitle.AccessibleName = "Title Combo Box Search";
            this.comboBoxTitle.AccessibleRole = System.Windows.Forms.AccessibleRole.ComboBox;
            this.comboBoxTitle.FormattingEnabled = true;
            this.comboBoxTitle.Items.AddRange(new object[] {
            "",
            "Dr",
            "Dr. PhD",
            "Father",
            "Master",
            "Miss",
            "Monsignor",
            "Mr",
            "Mrs",
            "Ms",
            "Not Known",
            "Professor",
            "Reverend",
            "Right Reverend",
            "Sister"});
            this.comboBoxTitle.Location = new System.Drawing.Point(296, 38);
            this.comboBoxTitle.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxTitle.Name = "comboBoxTitle";
            this.comboBoxTitle.Size = new System.Drawing.Size(75, 21);
            this.comboBoxTitle.TabIndex = 4;
            this.comboBoxTitle.SelectedIndexChanged += new System.EventHandler(this.comboBoxTitle_SelectedIndexChanged);
            // 
            // lblSearchSex
            // 
            this.lblSearchSex.AutoSize = true;
            this.lblSearchSex.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchSex.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchSex.Location = new System.Drawing.Point(260, 11);
            this.lblSearchSex.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchSex.Name = "lblSearchSex";
            this.lblSearchSex.Size = new System.Drawing.Size(25, 13);
            this.lblSearchSex.TabIndex = 48;
            this.lblSearchSex.Text = "Sex";
            // 
            // lblSearchLastname
            // 
            this.lblSearchLastname.AutoSize = true;
            this.lblSearchLastname.BackColor = System.Drawing.Color.Transparent;
            this.lblSearchLastname.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSearchLastname.Location = new System.Drawing.Point(54, 41);
            this.lblSearchLastname.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblSearchLastname.Name = "lblSearchLastname";
            this.lblSearchLastname.Size = new System.Drawing.Size(58, 13);
            this.lblSearchLastname.TabIndex = 47;
            this.lblSearchLastname.Text = "Last Name";
            // 
            // comboBoxSex
            // 
            this.comboBoxSex.AccessibleDescription = "Sex Combo Box Search";
            this.comboBoxSex.AccessibleName = "Sex Combo Box Search";
            this.comboBoxSex.AccessibleRole = System.Windows.Forms.AccessibleRole.ComboBox;
            this.comboBoxSex.FormattingEnabled = true;
            this.comboBoxSex.Items.AddRange(new object[] {
            "",
            "Male",
            "Female",
            "Not Known",
            "Not Specified"});
            this.comboBoxSex.Location = new System.Drawing.Point(296, 10);
            this.comboBoxSex.Margin = new System.Windows.Forms.Padding(2);
            this.comboBoxSex.Name = "comboBoxSex";
            this.comboBoxSex.Size = new System.Drawing.Size(75, 21);
            this.comboBoxSex.TabIndex = 3;
            this.comboBoxSex.SelectedIndexChanged += new System.EventHandler(this.comboBoxSex_SelectedIndexChanged);
            // 
            // btnClearSearch
            // 
            this.btnClearSearch.BackColor = System.Drawing.Color.LightSalmon;
            this.btnClearSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearSearch.ForeColor = System.Drawing.Color.White;
            this.btnClearSearch.Location = new System.Drawing.Point(658, 19);
            this.btnClearSearch.Margin = new System.Windows.Forms.Padding(1);
            this.btnClearSearch.Name = "btnClearSearch";
            this.btnClearSearch.Size = new System.Drawing.Size(64, 30);
            this.btnClearSearch.TabIndex = 7;
            this.btnClearSearch.Text = "Clear";
            this.btnClearSearch.UseVisualStyleBackColor = false;
            this.btnClearSearch.Click += new System.EventHandler(this.btnClearSearch_Click);
            // 
            // groupBoxGoogle
            // 
            this.groupBoxGoogle.Controls.Add(this.lblGoogleEmail);
            this.groupBoxGoogle.Controls.Add(this.txtboxGoogleEmail);
            this.groupBoxGoogle.Controls.Add(this.btnDeleteGoogleCalendar);
            this.groupBoxGoogle.Controls.Add(this.btnGoogleEmailClear);
            this.groupBoxGoogle.Controls.Add(this.btnGoogleEmailUpdate);
            this.groupBoxGoogle.Controls.Add(this.btnClearGoogleCalendar);
            this.groupBoxGoogle.Controls.Add(this.btnSampleGoogleData);
            this.groupBoxGoogle.Controls.Add(this.btnGoogleCreateShare);
            this.groupBoxGoogle.Location = new System.Drawing.Point(437, 344);
            this.groupBoxGoogle.Margin = new System.Windows.Forms.Padding(2);
            this.groupBoxGoogle.Name = "groupBoxGoogle";
            this.groupBoxGoogle.Padding = new System.Windows.Forms.Padding(2);
            this.groupBoxGoogle.Size = new System.Drawing.Size(199, 76);
            this.groupBoxGoogle.TabIndex = 45;
            this.groupBoxGoogle.TabStop = false;
            this.groupBoxGoogle.Text = "Google Calendar";
            // 
            // groupBoxNHSNet
            // 
            this.groupBoxNHSNet.Controls.Add(this.lblNHSNetEmail);
            this.groupBoxNHSNet.Controls.Add(this.txtboxNHSNetEmail);
            this.groupBoxNHSNet.Controls.Add(this.btnDeleteNHSNetCalendar);
            this.groupBoxNHSNet.Controls.Add(this.btnNHSNetEmailClear);
            this.groupBoxNHSNet.Controls.Add(this.btnNHSNetEmailUpdate);
            this.groupBoxNHSNet.Controls.Add(this.btnClearNHSNetCalendar);
            this.groupBoxNHSNet.Controls.Add(this.btnSampleNHSNetData);
            this.groupBoxNHSNet.Controls.Add(this.btnNHSNetCreateShare);
            this.groupBoxNHSNet.Location = new System.Drawing.Point(223, 344);
            this.groupBoxNHSNet.Margin = new System.Windows.Forms.Padding(2);
            this.groupBoxNHSNet.Name = "groupBoxNHSNet";
            this.groupBoxNHSNet.Padding = new System.Windows.Forms.Padding(2);
            this.groupBoxNHSNet.Size = new System.Drawing.Size(211, 76);
            this.groupBoxNHSNet.TabIndex = 46;
            this.groupBoxNHSNet.TabStop = false;
            this.groupBoxNHSNet.Text = "NHS Net Calendar";
            // 
            // lblNHSNetEmail
            // 
            this.lblNHSNetEmail.AutoSize = true;
            this.lblNHSNetEmail.BackColor = System.Drawing.Color.Transparent;
            this.lblNHSNetEmail.Enabled = false;
            this.lblNHSNetEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNHSNetEmail.Location = new System.Drawing.Point(1, 46);
            this.lblNHSNetEmail.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblNHSNetEmail.Name = "lblNHSNetEmail";
            this.lblNHSNetEmail.Size = new System.Drawing.Size(51, 13);
            this.lblNHSNetEmail.TabIndex = 0;
            this.lblNHSNetEmail.Text = "NHS mail";
            this.lblNHSNetEmail.Visible = false;
            // 
            // txtboxNHSNetEmail
            // 
            this.txtboxNHSNetEmail.Enabled = false;
            this.txtboxNHSNetEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxNHSNetEmail.Location = new System.Drawing.Point(57, 46);
            this.txtboxNHSNetEmail.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxNHSNetEmail.Name = "txtboxNHSNetEmail";
            this.txtboxNHSNetEmail.Size = new System.Drawing.Size(144, 20);
            this.txtboxNHSNetEmail.TabIndex = 16;
            this.txtboxNHSNetEmail.Visible = false;
            // 
            // btnDeleteNHSNetCalendar
            // 
            this.btnDeleteNHSNetCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnDeleteNHSNetCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDeleteNHSNetCalendar.BackgroundImage")));
            this.btnDeleteNHSNetCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDeleteNHSNetCalendar.Enabled = false;
            this.btnDeleteNHSNetCalendar.FlatAppearance.BorderSize = 0;
            this.btnDeleteNHSNetCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteNHSNetCalendar.ForeColor = System.Drawing.Color.White;
            this.btnDeleteNHSNetCalendar.Location = new System.Drawing.Point(172, 17);
            this.btnDeleteNHSNetCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnDeleteNHSNetCalendar.Name = "btnDeleteNHSNetCalendar";
            this.btnDeleteNHSNetCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnDeleteNHSNetCalendar.TabIndex = 22;
            this.btnDeleteNHSNetCalendar.UseVisualStyleBackColor = false;
            this.btnDeleteNHSNetCalendar.Visible = false;
            this.btnDeleteNHSNetCalendar.Click += new System.EventHandler(this.btnDeleteNHSNetCalendar_Click);
            // 
            // btnNHSNetEmailClear
            // 
            this.btnNHSNetEmailClear.BackColor = System.Drawing.Color.LightGray;
            this.btnNHSNetEmailClear.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnNHSNetEmailClear.BackgroundImage")));
            this.btnNHSNetEmailClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnNHSNetEmailClear.Enabled = false;
            this.btnNHSNetEmailClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNHSNetEmailClear.ForeColor = System.Drawing.Color.White;
            this.btnNHSNetEmailClear.Location = new System.Drawing.Point(36, 15);
            this.btnNHSNetEmailClear.Margin = new System.Windows.Forms.Padding(1);
            this.btnNHSNetEmailClear.Name = "btnNHSNetEmailClear";
            this.btnNHSNetEmailClear.Size = new System.Drawing.Size(29, 27);
            this.btnNHSNetEmailClear.TabIndex = 18;
            this.btnNHSNetEmailClear.UseVisualStyleBackColor = false;
            this.btnNHSNetEmailClear.Visible = false;
            this.btnNHSNetEmailClear.Click += new System.EventHandler(this.btnNHSNetEmailClear_Click);
            // 
            // btnNHSNetEmailUpdate
            // 
            this.btnNHSNetEmailUpdate.BackColor = System.Drawing.Color.LightGray;
            this.btnNHSNetEmailUpdate.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnNHSNetEmailUpdate.BackgroundImage")));
            this.btnNHSNetEmailUpdate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnNHSNetEmailUpdate.Enabled = false;
            this.btnNHSNetEmailUpdate.FlatAppearance.BorderSize = 0;
            this.btnNHSNetEmailUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNHSNetEmailUpdate.ForeColor = System.Drawing.Color.White;
            this.btnNHSNetEmailUpdate.Location = new System.Drawing.Point(6, 15);
            this.btnNHSNetEmailUpdate.Margin = new System.Windows.Forms.Padding(1);
            this.btnNHSNetEmailUpdate.Name = "btnNHSNetEmailUpdate";
            this.btnNHSNetEmailUpdate.Size = new System.Drawing.Size(29, 27);
            this.btnNHSNetEmailUpdate.TabIndex = 17;
            this.btnNHSNetEmailUpdate.UseVisualStyleBackColor = false;
            this.btnNHSNetEmailUpdate.Visible = false;
            this.btnNHSNetEmailUpdate.Click += new System.EventHandler(this.btnNHSNetEmailUpdate_Click);
            // 
            // btnClearNHSNetCalendar
            // 
            this.btnClearNHSNetCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnClearNHSNetCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClearNHSNetCalendar.BackgroundImage")));
            this.btnClearNHSNetCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClearNHSNetCalendar.Enabled = false;
            this.btnClearNHSNetCalendar.FlatAppearance.BorderSize = 0;
            this.btnClearNHSNetCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearNHSNetCalendar.ForeColor = System.Drawing.Color.White;
            this.btnClearNHSNetCalendar.Location = new System.Drawing.Point(140, 16);
            this.btnClearNHSNetCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnClearNHSNetCalendar.Name = "btnClearNHSNetCalendar";
            this.btnClearNHSNetCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnClearNHSNetCalendar.TabIndex = 21;
            this.btnClearNHSNetCalendar.UseVisualStyleBackColor = false;
            this.btnClearNHSNetCalendar.Visible = false;
            this.btnClearNHSNetCalendar.Click += new System.EventHandler(this.btnClearNHSNetCalendar_Click);
            // 
            // btnSampleNHSNetData
            // 
            this.btnSampleNHSNetData.BackColor = System.Drawing.Color.LightGray;
            this.btnSampleNHSNetData.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSampleNHSNetData.BackgroundImage")));
            this.btnSampleNHSNetData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSampleNHSNetData.Enabled = false;
            this.btnSampleNHSNetData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSampleNHSNetData.ForeColor = System.Drawing.Color.White;
            this.btnSampleNHSNetData.Location = new System.Drawing.Point(109, 17);
            this.btnSampleNHSNetData.Margin = new System.Windows.Forms.Padding(1);
            this.btnSampleNHSNetData.Name = "btnSampleNHSNetData";
            this.btnSampleNHSNetData.Size = new System.Drawing.Size(29, 27);
            this.btnSampleNHSNetData.TabIndex = 20;
            this.btnSampleNHSNetData.UseVisualStyleBackColor = false;
            this.btnSampleNHSNetData.Visible = false;
            this.btnSampleNHSNetData.Click += new System.EventHandler(this.btnSampleNHSNetData_Click);
            // 
            // btnNHSNetCreateShare
            // 
            this.btnNHSNetCreateShare.BackColor = System.Drawing.Color.LightGray;
            this.btnNHSNetCreateShare.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnNHSNetCreateShare.BackgroundImage")));
            this.btnNHSNetCreateShare.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnNHSNetCreateShare.Enabled = false;
            this.btnNHSNetCreateShare.FlatAppearance.BorderSize = 0;
            this.btnNHSNetCreateShare.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNHSNetCreateShare.ForeColor = System.Drawing.Color.White;
            this.btnNHSNetCreateShare.Location = new System.Drawing.Point(68, 17);
            this.btnNHSNetCreateShare.Margin = new System.Windows.Forms.Padding(1);
            this.btnNHSNetCreateShare.Name = "btnNHSNetCreateShare";
            this.btnNHSNetCreateShare.Size = new System.Drawing.Size(39, 27);
            this.btnNHSNetCreateShare.TabIndex = 19;
            this.btnNHSNetCreateShare.UseVisualStyleBackColor = false;
            this.btnNHSNetCreateShare.Visible = false;
            this.btnNHSNetCreateShare.Click += new System.EventHandler(this.btnNHSNetCreateShare_Click);
            // 
            // groupBoxExchange
            // 
            this.groupBoxExchange.Controls.Add(this.lblExchangeEmail);
            this.groupBoxExchange.Controls.Add(this.txtboxExchangeEmail);
            this.groupBoxExchange.Controls.Add(this.btnDeleteExchangeCalendar);
            this.groupBoxExchange.Controls.Add(this.btnExchangeEmailClear);
            this.groupBoxExchange.Controls.Add(this.btnExchangeEmailUpdate);
            this.groupBoxExchange.Controls.Add(this.btnClearExchangeCalendar);
            this.groupBoxExchange.Controls.Add(this.btnSampleExchangeData);
            this.groupBoxExchange.Controls.Add(this.btnExchangeCreateShare);
            this.groupBoxExchange.Location = new System.Drawing.Point(9, 344);
            this.groupBoxExchange.Margin = new System.Windows.Forms.Padding(2);
            this.groupBoxExchange.Name = "groupBoxExchange";
            this.groupBoxExchange.Padding = new System.Windows.Forms.Padding(2);
            this.groupBoxExchange.Size = new System.Drawing.Size(211, 76);
            this.groupBoxExchange.TabIndex = 47;
            this.groupBoxExchange.TabStop = false;
            this.groupBoxExchange.Text = "Exchange Calendar";
            // 
            // lblExchangeEmail
            // 
            this.lblExchangeEmail.AutoSize = true;
            this.lblExchangeEmail.BackColor = System.Drawing.Color.Transparent;
            this.lblExchangeEmail.Enabled = false;
            this.lblExchangeEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblExchangeEmail.Location = new System.Drawing.Point(1, 46);
            this.lblExchangeEmail.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.lblExchangeEmail.Name = "lblExchangeEmail";
            this.lblExchangeEmail.Size = new System.Drawing.Size(76, 13);
            this.lblExchangeEmail.TabIndex = 0;
            this.lblExchangeEmail.Text = "Exchange mail";
            this.lblExchangeEmail.Visible = false;
            // 
            // txtboxExchangeEmail
            // 
            this.txtboxExchangeEmail.Enabled = false;
            this.txtboxExchangeEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxExchangeEmail.Location = new System.Drawing.Point(80, 46);
            this.txtboxExchangeEmail.Margin = new System.Windows.Forms.Padding(1);
            this.txtboxExchangeEmail.Name = "txtboxExchangeEmail";
            this.txtboxExchangeEmail.Size = new System.Drawing.Size(130, 20);
            this.txtboxExchangeEmail.TabIndex = 9;
            this.txtboxExchangeEmail.Visible = false;
            // 
            // btnDeleteExchangeCalendar
            // 
            this.btnDeleteExchangeCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnDeleteExchangeCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnDeleteExchangeCalendar.BackgroundImage")));
            this.btnDeleteExchangeCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnDeleteExchangeCalendar.Enabled = false;
            this.btnDeleteExchangeCalendar.FlatAppearance.BorderSize = 0;
            this.btnDeleteExchangeCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteExchangeCalendar.ForeColor = System.Drawing.Color.White;
            this.btnDeleteExchangeCalendar.Location = new System.Drawing.Point(162, 17);
            this.btnDeleteExchangeCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnDeleteExchangeCalendar.Name = "btnDeleteExchangeCalendar";
            this.btnDeleteExchangeCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnDeleteExchangeCalendar.TabIndex = 15;
            this.btnDeleteExchangeCalendar.UseVisualStyleBackColor = false;
            this.btnDeleteExchangeCalendar.Visible = false;
            this.btnDeleteExchangeCalendar.Click += new System.EventHandler(this.btnDeleteExchangeCalendar_Click);
            // 
            // btnExchangeEmailClear
            // 
            this.btnExchangeEmailClear.BackColor = System.Drawing.Color.LightGray;
            this.btnExchangeEmailClear.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExchangeEmailClear.BackgroundImage")));
            this.btnExchangeEmailClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnExchangeEmailClear.Enabled = false;
            this.btnExchangeEmailClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeEmailClear.ForeColor = System.Drawing.Color.White;
            this.btnExchangeEmailClear.Location = new System.Drawing.Point(36, 15);
            this.btnExchangeEmailClear.Margin = new System.Windows.Forms.Padding(1);
            this.btnExchangeEmailClear.Name = "btnExchangeEmailClear";
            this.btnExchangeEmailClear.Size = new System.Drawing.Size(29, 27);
            this.btnExchangeEmailClear.TabIndex = 11;
            this.btnExchangeEmailClear.UseVisualStyleBackColor = false;
            this.btnExchangeEmailClear.Visible = false;
            this.btnExchangeEmailClear.Click += new System.EventHandler(this.btnExchangeEmailClear_Click);
            // 
            // btnExchangeEmailUpdate
            // 
            this.btnExchangeEmailUpdate.BackColor = System.Drawing.Color.LightGray;
            this.btnExchangeEmailUpdate.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExchangeEmailUpdate.BackgroundImage")));
            this.btnExchangeEmailUpdate.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnExchangeEmailUpdate.Enabled = false;
            this.btnExchangeEmailUpdate.FlatAppearance.BorderSize = 0;
            this.btnExchangeEmailUpdate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.900001F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeEmailUpdate.ForeColor = System.Drawing.Color.White;
            this.btnExchangeEmailUpdate.Location = new System.Drawing.Point(6, 15);
            this.btnExchangeEmailUpdate.Margin = new System.Windows.Forms.Padding(1);
            this.btnExchangeEmailUpdate.Name = "btnExchangeEmailUpdate";
            this.btnExchangeEmailUpdate.Size = new System.Drawing.Size(29, 27);
            this.btnExchangeEmailUpdate.TabIndex = 10;
            this.btnExchangeEmailUpdate.UseVisualStyleBackColor = false;
            this.btnExchangeEmailUpdate.Visible = false;
            this.btnExchangeEmailUpdate.Click += new System.EventHandler(this.btnExchangeEmailUpdate_Click);
            // 
            // btnClearExchangeCalendar
            // 
            this.btnClearExchangeCalendar.BackColor = System.Drawing.Color.LightGray;
            this.btnClearExchangeCalendar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnClearExchangeCalendar.BackgroundImage")));
            this.btnClearExchangeCalendar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnClearExchangeCalendar.Enabled = false;
            this.btnClearExchangeCalendar.FlatAppearance.BorderSize = 0;
            this.btnClearExchangeCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearExchangeCalendar.ForeColor = System.Drawing.Color.White;
            this.btnClearExchangeCalendar.Location = new System.Drawing.Point(131, 17);
            this.btnClearExchangeCalendar.Margin = new System.Windows.Forms.Padding(1);
            this.btnClearExchangeCalendar.Name = "btnClearExchangeCalendar";
            this.btnClearExchangeCalendar.Size = new System.Drawing.Size(29, 27);
            this.btnClearExchangeCalendar.TabIndex = 14;
            this.btnClearExchangeCalendar.UseVisualStyleBackColor = false;
            this.btnClearExchangeCalendar.Visible = false;
            this.btnClearExchangeCalendar.Click += new System.EventHandler(this.btnClearExchangeCalendar_Click);
            // 
            // btnSampleExchangeData
            // 
            this.btnSampleExchangeData.BackColor = System.Drawing.Color.LightGray;
            this.btnSampleExchangeData.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSampleExchangeData.BackgroundImage")));
            this.btnSampleExchangeData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSampleExchangeData.Enabled = false;
            this.btnSampleExchangeData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSampleExchangeData.ForeColor = System.Drawing.Color.White;
            this.btnSampleExchangeData.Location = new System.Drawing.Point(100, 17);
            this.btnSampleExchangeData.Margin = new System.Windows.Forms.Padding(1);
            this.btnSampleExchangeData.Name = "btnSampleExchangeData";
            this.btnSampleExchangeData.Size = new System.Drawing.Size(29, 27);
            this.btnSampleExchangeData.TabIndex = 13;
            this.btnSampleExchangeData.UseVisualStyleBackColor = false;
            this.btnSampleExchangeData.Visible = false;
            this.btnSampleExchangeData.Click += new System.EventHandler(this.btnSampleExchangeData_Click);
            // 
            // btnExchangeCreateShare
            // 
            this.btnExchangeCreateShare.BackColor = System.Drawing.Color.LightGray;
            this.btnExchangeCreateShare.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExchangeCreateShare.BackgroundImage")));
            this.btnExchangeCreateShare.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnExchangeCreateShare.Enabled = false;
            this.btnExchangeCreateShare.FlatAppearance.BorderSize = 0;
            this.btnExchangeCreateShare.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExchangeCreateShare.ForeColor = System.Drawing.Color.White;
            this.btnExchangeCreateShare.Location = new System.Drawing.Point(68, 17);
            this.btnExchangeCreateShare.Margin = new System.Windows.Forms.Padding(1);
            this.btnExchangeCreateShare.Name = "btnExchangeCreateShare";
            this.btnExchangeCreateShare.Size = new System.Drawing.Size(29, 27);
            this.btnExchangeCreateShare.TabIndex = 12;
            this.btnExchangeCreateShare.UseVisualStyleBackColor = false;
            this.btnExchangeCreateShare.Visible = false;
            this.btnExchangeCreateShare.Click += new System.EventHandler(this.btnExchangeCreateShare_Click);
            // 
            // HealthCalendarAdmin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(789, 478);
            this.Controls.Add(this.groupBoxExchange);
            this.Controls.Add(this.groupBoxNHSNet);
            this.Controls.Add(this.groupBoxSearch);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.dgvSubscribers);
            this.Controls.Add(this.pnlFooter);
            this.Controls.Add(this.groupBoxGoogle);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.1F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(1);
            this.MaximumSize = new System.Drawing.Size(1205, 611);
            this.MinimumSize = new System.Drawing.Size(805, 351);
            this.Name = "HealthCalendarAdmin";
            this.Text = "Health Calendar Admin";
            this.Load += new System.EventHandler(this.HealthCalendarAdmin_Load);
            this.Shown += new System.EventHandler(this.HealthCalendarAdmin_Shown);
            this.Leave += new System.EventHandler(this.HealthCalendarAdmin_Leave);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubscribers)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.pnlFooter.ResumeLayout(false);
            this.pnlFooter.PerformLayout();
            this.groupBoxSearch.ResumeLayout(false);
            this.groupBoxSearch.PerformLayout();
            this.groupBoxGoogle.ResumeLayout(false);
            this.groupBoxGoogle.PerformLayout();
            this.groupBoxNHSNet.ResumeLayout(false);
            this.groupBoxNHSNet.PerformLayout();
            this.groupBoxExchange.ResumeLayout(false);
            this.groupBoxExchange.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtboxSearchFirstname;
        private System.Windows.Forms.Label lblSearchFirstname;
        private System.Windows.Forms.DataGridView dgvSubscribers;
        private System.Windows.Forms.Button btnGoogleEmailClear;
        private System.Windows.Forms.Button btnGoogleEmailUpdate;
        private System.Windows.Forms.TextBox txtboxGoogleEmail;
        private System.Windows.Forms.Label lblGoogleEmail;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel pnlFooter;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btnGoogleCreateShare;
        private System.Windows.Forms.Button btnSampleGoogleData;
        private System.Windows.Forms.Button btnClearGoogleCalendar;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.ComponentModel.BackgroundWorker bgWorker;
        private System.Windows.Forms.Label lblPercent;
        private System.Windows.Forms.Button btnDeleteGoogleCalendar;
        private System.Windows.Forms.TextBox txtboxSearchLastname;
        private System.Windows.Forms.GroupBox groupBoxSearch;
        private System.Windows.Forms.Button btnClearSearch;
        private System.Windows.Forms.ComboBox comboBoxSex;
        private System.Windows.Forms.Label lblSearchLastname;
        private System.Windows.Forms.Label lblSearchSex;
        private System.Windows.Forms.Label lblSearchTitle;
        private System.Windows.Forms.ComboBox comboBoxTitle;
        private System.Windows.Forms.Label lblSearchMainIdentifier;
        private System.Windows.Forms.TextBox txtboxSearchMainIdentifier;
        private System.Windows.Forms.Label lblSearchOccupation;
        private System.Windows.Forms.ComboBox comboBoxOccupation;
        private System.Windows.Forms.GroupBox groupBoxGoogle;
        private System.Windows.Forms.GroupBox groupBoxNHSNet;
        private System.Windows.Forms.Label lblNHSNetEmail;
        private System.Windows.Forms.TextBox txtboxNHSNetEmail;
        private System.Windows.Forms.Button btnDeleteNHSNetCalendar;
        private System.Windows.Forms.Button btnNHSNetEmailClear;
        private System.Windows.Forms.Button btnNHSNetEmailUpdate;
        private System.Windows.Forms.Button btnClearNHSNetCalendar;
        private System.Windows.Forms.Button btnSampleNHSNetData;
        private System.Windows.Forms.Button btnNHSNetCreateShare;
        private System.Windows.Forms.GroupBox groupBoxExchange;
        private System.Windows.Forms.Label lblExchangeEmail;
        private System.Windows.Forms.TextBox txtboxExchangeEmail;
        private System.Windows.Forms.Button btnDeleteExchangeCalendar;
        private System.Windows.Forms.Button btnExchangeEmailClear;
        private System.Windows.Forms.Button btnExchangeEmailUpdate;
        private System.Windows.Forms.Button btnClearExchangeCalendar;
        private System.Windows.Forms.Button btnSampleExchangeData;
        private System.Windows.Forms.Button btnExchangeCreateShare;
    }
}

