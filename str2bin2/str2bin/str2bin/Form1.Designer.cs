﻿namespace str2bin
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.fill = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.baseAddr = new System.Windows.Forms.TextBox();
            this.upAddr = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.allEnter = new System.Windows.Forms.CheckBox();
            this.allnewLine = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "空白填充：  0x";
            // 
            // fill
            // 
            this.fill.FormattingEnabled = true;
            this.fill.Items.AddRange(new object[] {
            "FF",
            "00"});
            this.fill.Location = new System.Drawing.Point(108, 9);
            this.fill.Name = "fill";
            this.fill.Size = new System.Drawing.Size(60, 20);
            this.fill.TabIndex = 1;
            this.fill.Text = "FF";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(184, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "起始地址：  0x";
            // 
            // baseAddr
            // 
            this.baseAddr.Location = new System.Drawing.Point(279, 9);
            this.baseAddr.Name = "baseAddr";
            this.baseAddr.Size = new System.Drawing.Size(59, 21);
            this.baseAddr.TabIndex = 3;
            this.baseAddr.Text = "0";
            this.baseAddr.TextChanged += new System.EventHandler(this.baseAddr_TextChanged);
            // 
            // upAddr
            // 
            this.upAddr.Location = new System.Drawing.Point(440, 10);
            this.upAddr.Name = "upAddr";
            this.upAddr.Size = new System.Drawing.Size(66, 21);
            this.upAddr.TabIndex = 3;
            this.upAddr.Text = "100";
            this.upAddr.TextChanged += new System.EventHandler(this.upAddr_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(345, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "自增地址：  0x";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(67, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 4;
            this.label4.Text = "地址";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(137, 87);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 5;
            this.label5.Text = "字符串";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(394, 87);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "回车";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 45);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(145, 24);
            this.button1.TabIndex = 11;
            this.button1.Text = "保存";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(187, 45);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(145, 24);
            this.button2.TabIndex = 12;
            this.button2.Text = "打开";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(359, 45);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(145, 24);
            this.button3.TabIndex = 13;
            this.button3.Text = "写入";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(13, 87);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(29, 12);
            this.label9.TabIndex = 14;
            this.label9.Text = "序号";
            // 
            // allEnter
            // 
            this.allEnter.AutoSize = true;
            this.allEnter.Location = new System.Drawing.Point(426, 86);
            this.allEnter.Name = "allEnter";
            this.allEnter.Size = new System.Drawing.Size(15, 14);
            this.allEnter.TabIndex = 15;
            this.allEnter.UseVisualStyleBackColor = true;
            this.allEnter.CheckedChanged += new System.EventHandler(this.allEnter_CheckedChanged);
            // 
            // allnewLine
            // 
            this.allnewLine.AutoSize = true;
            this.allnewLine.Location = new System.Drawing.Point(477, 86);
            this.allnewLine.Name = "allnewLine";
            this.allnewLine.Size = new System.Drawing.Size(15, 14);
            this.allnewLine.TabIndex = 16;
            this.allnewLine.UseVisualStyleBackColor = true;
            this.allnewLine.CheckedChanged += new System.EventHandler(this.allnewLine_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(447, 87);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "换行";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(525, 618);
            this.Controls.Add(this.allnewLine);
            this.Controls.Add(this.allEnter);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.upAddr);
            this.Controls.Add(this.baseAddr);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.fill);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(541, 656);
            this.MinimumSize = new System.Drawing.Size(541, 656);
            this.Name = "Form1";
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox fill;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox baseAddr;
        private System.Windows.Forms.TextBox upAddr;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox allEnter;
        private System.Windows.Forms.CheckBox allnewLine;
        private System.Windows.Forms.Label label7;




    }
}

