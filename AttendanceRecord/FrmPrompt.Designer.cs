﻿namespace AttendanceRecord
{
    partial class FrmPrompt
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
            this.timer_Close_Excel = new System.Windows.Forms.Timer(this.components);
            this.btnCloseAllExcel = new System.Windows.Forms.Button();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // timer_Close_Excel
            // 
            this.timer_Close_Excel.Interval = 1000;
            this.timer_Close_Excel.Tick += new System.EventHandler(this.timer_Close_Excel_Tick);
            // 
            // btnCloseAllExcel
            // 
            this.btnCloseAllExcel.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCloseAllExcel.Location = new System.Drawing.Point(517, 202);
            this.btnCloseAllExcel.Name = "btnCloseAllExcel";
            this.btnCloseAllExcel.Size = new System.Drawing.Size(128, 45);
            this.btnCloseAllExcel.TabIndex = 0;
            this.btnCloseAllExcel.Text = "20";
            this.btnCloseAllExcel.UseVisualStyleBackColor = true;
            // 
            // lblPrompt
            // 
            this.lblPrompt.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblPrompt.ForeColor = System.Drawing.Color.Black;
            this.lblPrompt.Location = new System.Drawing.Point(18, 202);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(479, 44);
            this.lblPrompt.TabIndex = 1;
            this.lblPrompt.Text = "请先保存并关闭已经打开的Excel";
            this.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // FrmPrompt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(657, 259);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.btnCloseAllExcel);
            this.Name = "FrmPrompt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "提示：";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmPrompt_FormClosed);
            this.Load += new System.EventHandler(this.FrmPrompt_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timer_Close_Excel;
        private System.Windows.Forms.Button btnCloseAllExcel;
        private System.Windows.Forms.Label lblPrompt;
    }
}