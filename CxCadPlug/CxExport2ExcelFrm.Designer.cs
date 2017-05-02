namespace CxCadPlug
{
    partial class export2ExcelFrm
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
            this.layerChb = new System.Windows.Forms.CheckedListBox();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.okBtn = new System.Windows.Forms.Button();
            this.selectBtn = new System.Windows.Forms.Button();
            this.exportTxb = new System.Windows.Forms.TextBox();
            this.exportLab = new System.Windows.Forms.Label();
            this.statusPgBar = new System.Windows.Forms.ProgressBar();
            this.statusLabel = new System.Windows.Forms.Label();
            this.inOptionGpb = new System.Windows.Forms.GroupBox();
            this.decimalSpin = new System.Windows.Forms.NumericUpDown();
            this.decimalLab = new System.Windows.Forms.Label();
            this.compareCmb = new System.Windows.Forms.ComboBox();
            this.CompareLab = new System.Windows.Forms.Label();
            this.typeCmb = new System.Windows.Forms.ComboBox();
            this.typeLab = new System.Windows.Forms.Label();
            this.inOptionGpb.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.decimalSpin)).BeginInit();
            this.SuspendLayout();
            // 
            // layerChb
            // 
            this.layerChb.CheckOnClick = true;
            this.layerChb.FormattingEnabled = true;
            this.layerChb.Location = new System.Drawing.Point(9, 130);
            this.layerChb.Name = "layerChb";
            this.layerChb.Size = new System.Drawing.Size(377, 212);
            this.layerChb.TabIndex = 19;
            // 
            // cancelBtn
            // 
            this.cancelBtn.Location = new System.Drawing.Point(318, 396);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(75, 23);
            this.cancelBtn.TabIndex = 18;
            this.cancelBtn.Text = "取消";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // okBtn
            // 
            this.okBtn.Location = new System.Drawing.Point(237, 396);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(75, 23);
            this.okBtn.TabIndex = 17;
            this.okBtn.Text = "确定";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // selectBtn
            // 
            this.selectBtn.Location = new System.Drawing.Point(318, 346);
            this.selectBtn.Name = "selectBtn";
            this.selectBtn.Size = new System.Drawing.Size(75, 23);
            this.selectBtn.TabIndex = 16;
            this.selectBtn.Text = "浏览";
            this.selectBtn.UseVisualStyleBackColor = true;
            this.selectBtn.Click += new System.EventHandler(this.selectBtn_Click);
            // 
            // exportTxb
            // 
            this.exportTxb.Location = new System.Drawing.Point(70, 348);
            this.exportTxb.Name = "exportTxb";
            this.exportTxb.Size = new System.Drawing.Size(242, 21);
            this.exportTxb.TabIndex = 15;
            // 
            // exportLab
            // 
            this.exportLab.AutoSize = true;
            this.exportLab.Location = new System.Drawing.Point(8, 351);
            this.exportLab.Name = "exportLab";
            this.exportLab.Size = new System.Drawing.Size(65, 12);
            this.exportLab.TabIndex = 14;
            this.exportLab.Text = "导出路径：";
            // 
            // statusPgBar
            // 
            this.statusPgBar.Location = new System.Drawing.Point(12, 398);
            this.statusPgBar.Name = "statusPgBar";
            this.statusPgBar.Size = new System.Drawing.Size(220, 19);
            this.statusPgBar.TabIndex = 20;
            this.statusPgBar.Tag = "";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(8, 378);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(41, 12);
            this.statusLabel.TabIndex = 21;
            this.statusLabel.Text = "已完成";
            // 
            // inOptionGpb
            // 
            this.inOptionGpb.Controls.Add(this.decimalSpin);
            this.inOptionGpb.Controls.Add(this.decimalLab);
            this.inOptionGpb.Controls.Add(this.compareCmb);
            this.inOptionGpb.Controls.Add(this.CompareLab);
            this.inOptionGpb.Controls.Add(this.typeCmb);
            this.inOptionGpb.Controls.Add(this.typeLab);
            this.inOptionGpb.Location = new System.Drawing.Point(9, 12);
            this.inOptionGpb.Name = "inOptionGpb";
            this.inOptionGpb.Size = new System.Drawing.Size(377, 112);
            this.inOptionGpb.TabIndex = 22;
            this.inOptionGpb.TabStop = false;
            this.inOptionGpb.Text = "输入参数设置";
            // 
            // decimalSpin
            // 
            this.decimalSpin.Location = new System.Drawing.Point(95, 79);
            this.decimalSpin.Maximum = new decimal(new int[] {
            11,
            0,
            0,
            0});
            this.decimalSpin.Name = "decimalSpin";
            this.decimalSpin.Size = new System.Drawing.Size(65, 21);
            this.decimalSpin.TabIndex = 19;
            this.decimalSpin.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // decimalLab
            // 
            this.decimalLab.AutoSize = true;
            this.decimalLab.Location = new System.Drawing.Point(6, 84);
            this.decimalLab.Name = "decimalLab";
            this.decimalLab.Size = new System.Drawing.Size(65, 12);
            this.decimalLab.TabIndex = 18;
            this.decimalLab.Text = "小数位数：";
            // 
            // compareCmb
            // 
            this.compareCmb.FormattingEnabled = true;
            this.compareCmb.Location = new System.Drawing.Point(95, 52);
            this.compareCmb.Name = "compareCmb";
            this.compareCmb.Size = new System.Drawing.Size(272, 20);
            this.compareCmb.TabIndex = 17;
            // 
            // CompareLab
            // 
            this.CompareLab.AutoSize = true;
            this.CompareLab.Location = new System.Drawing.Point(6, 55);
            this.CompareLab.Name = "CompareLab";
            this.CompareLab.Size = new System.Drawing.Size(77, 12);
            this.CompareLab.TabIndex = 16;
            this.CompareLab.Text = "符号对照表：";
            // 
            // typeCmb
            // 
            this.typeCmb.FormattingEnabled = true;
            this.typeCmb.Location = new System.Drawing.Point(95, 20);
            this.typeCmb.Name = "typeCmb";
            this.typeCmb.Size = new System.Drawing.Size(272, 20);
            this.typeCmb.TabIndex = 15;
            // 
            // typeLab
            // 
            this.typeLab.AutoSize = true;
            this.typeLab.Location = new System.Drawing.Point(6, 28);
            this.typeLab.Name = "typeLab";
            this.typeLab.Size = new System.Drawing.Size(65, 12);
            this.typeLab.TabIndex = 14;
            this.typeLab.Text = "管网类型：";
            // 
            // export2ExcelFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 426);
            this.Controls.Add(this.inOptionGpb);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.statusPgBar);
            this.Controls.Add(this.layerChb);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.okBtn);
            this.Controls.Add(this.selectBtn);
            this.Controls.Add(this.exportTxb);
            this.Controls.Add(this.exportLab);
            this.Name = "export2ExcelFrm";
            this.Text = "导出为Excel";
            this.inOptionGpb.ResumeLayout(false);
            this.inOptionGpb.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.decimalSpin)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox layerChb;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button okBtn;
        private System.Windows.Forms.Button selectBtn;
        private System.Windows.Forms.TextBox exportTxb;
        private System.Windows.Forms.Label exportLab;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.GroupBox inOptionGpb;
        private System.Windows.Forms.Label decimalLab;
        private System.Windows.Forms.ComboBox compareCmb;
        private System.Windows.Forms.Label CompareLab;
        private System.Windows.Forms.ComboBox typeCmb;
        private System.Windows.Forms.Label typeLab;
        private System.Windows.Forms.NumericUpDown decimalSpin;
        private System.Windows.Forms.ProgressBar statusPgBar;
    }
}