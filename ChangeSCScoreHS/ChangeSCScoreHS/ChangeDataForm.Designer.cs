namespace ChangeSCScoreHS
{
    partial class ChangeDataForm
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
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.cbxTotal = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.cbxSScore = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbxAScore = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnRun = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // iptSchoolYear
            // 
            this.iptSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSchoolYear.Location = new System.Drawing.Point(116, 21);
            this.iptSchoolYear.MaxValue = 3000;
            this.iptSchoolYear.MinValue = 0;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.ShowUpDown = true;
            this.iptSchoolYear.Size = new System.Drawing.Size(80, 25);
            this.iptSchoolYear.TabIndex = 0;
            // 
            // iptSemester
            // 
            this.iptSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSemester.Location = new System.Drawing.Point(259, 21);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.ShowUpDown = true;
            this.iptSemester.Size = new System.Drawing.Size(66, 25);
            this.iptSemester.TabIndex = 1;
            this.iptSemester.Value = 1;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(63, 23);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 2;
            this.labelX1.Text = "學年度";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(219, 23);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 21);
            this.labelX2.TabIndex = 3;
            this.labelX2.Text = "學期";
            // 
            // cbxTotal
            // 
            this.cbxTotal.DisplayMember = "Text";
            this.cbxTotal.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxTotal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxTotal.FormattingEnabled = true;
            this.cbxTotal.ItemHeight = 19;
            this.cbxTotal.Location = new System.Drawing.Point(116, 71);
            this.cbxTotal.Name = "cbxTotal";
            this.cbxTotal.Size = new System.Drawing.Size(209, 25);
            this.cbxTotal.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxTotal.TabIndex = 4;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(23, 73);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(87, 21);
            this.labelX3.TabIndex = 5;
            this.labelX3.Text = "合併評量名稱";
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            this.labelX4.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(23, 119);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(87, 21);
            this.labelX4.TabIndex = 6;
            this.labelX4.Text = "來源定期名稱";
            // 
            // labelX5
            // 
            this.labelX5.AutoSize = true;
            this.labelX5.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.Class = "";
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Location = new System.Drawing.Point(23, 162);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(87, 21);
            this.labelX5.TabIndex = 7;
            this.labelX5.Text = "來源平時名稱";
            // 
            // cbxSScore
            // 
            this.cbxSScore.DisplayMember = "Text";
            this.cbxSScore.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxSScore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSScore.FormattingEnabled = true;
            this.cbxSScore.ItemHeight = 19;
            this.cbxSScore.Location = new System.Drawing.Point(116, 115);
            this.cbxSScore.Name = "cbxSScore";
            this.cbxSScore.Size = new System.Drawing.Size(209, 25);
            this.cbxSScore.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxSScore.TabIndex = 8;
            // 
            // cbxAScore
            // 
            this.cbxAScore.DisplayMember = "Text";
            this.cbxAScore.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbxAScore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxAScore.FormattingEnabled = true;
            this.cbxAScore.ItemHeight = 19;
            this.cbxAScore.Location = new System.Drawing.Point(116, 158);
            this.cbxAScore.Name = "cbxAScore";
            this.cbxAScore.Size = new System.Drawing.Size(209, 25);
            this.cbxAScore.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbxAScore.TabIndex = 9;
            // 
            // btnRun
            // 
            this.btnRun.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnRun.AutoSize = true;
            this.btnRun.BackColor = System.Drawing.Color.Transparent;
            this.btnRun.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnRun.Location = new System.Drawing.Point(162, 203);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 25);
            this.btnRun.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnRun.TabIndex = 10;
            this.btnRun.Text = "轉換";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(250, 203);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 11;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // ChangeDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(349, 242);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.cbxAScore);
            this.Controls.Add(this.cbxSScore);
            this.Controls.Add(this.labelX5);
            this.Controls.Add(this.labelX4);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.cbxTotal);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.iptSchoolYear);
            this.DoubleBuffered = true;
            this.Name = "ChangeDataForm";
            this.Text = "合併定期與評量成績並轉換小考成績(新竹格式)";
            this.Load += new System.EventHandler(this.ChangeDataForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxTotal;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.LabelX labelX5;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxSScore;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbxAScore;
        private DevComponents.DotNetBar.ButtonX btnRun;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}