namespace xlsScheoutput
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkHi = new System.Windows.Forms.CheckBox();
            this.chkFumei = new System.Windows.Forms.CheckBox();
            this.chkKyu = new System.Windows.Forms.CheckBox();
            this.chkTai = new System.Windows.Forms.CheckBox();
            this.chkGen = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkHi);
            this.groupBox1.Controls.Add(this.chkFumei);
            this.groupBox1.Controls.Add(this.chkKyu);
            this.groupBox1.Controls.Add(this.chkTai);
            this.groupBox1.Controls.Add(this.chkGen);
            this.groupBox1.Location = new System.Drawing.Point(12, 22);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(350, 64);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "発行する会員歴";
            // 
            // chkHi
            // 
            this.chkHi.AutoSize = true;
            this.chkHi.Location = new System.Drawing.Point(283, 28);
            this.chkHi.Name = "chkHi";
            this.chkHi.Size = new System.Drawing.Size(43, 23);
            this.chkHi.TabIndex = 4;
            this.chkHi.Text = "非";
            this.chkHi.UseVisualStyleBackColor = true;
            // 
            // chkFumei
            // 
            this.chkFumei.AutoSize = true;
            this.chkFumei.Location = new System.Drawing.Point(208, 28);
            this.chkFumei.Name = "chkFumei";
            this.chkFumei.Size = new System.Drawing.Size(58, 23);
            this.chkFumei.TabIndex = 3;
            this.chkFumei.Text = "不明";
            this.chkFumei.UseVisualStyleBackColor = true;
            // 
            // chkKyu
            // 
            this.chkKyu.AutoSize = true;
            this.chkKyu.Location = new System.Drawing.Point(148, 28);
            this.chkKyu.Name = "chkKyu";
            this.chkKyu.Size = new System.Drawing.Size(43, 23);
            this.chkKyu.TabIndex = 2;
            this.chkKyu.Text = "休";
            this.chkKyu.UseVisualStyleBackColor = true;
            // 
            // chkTai
            // 
            this.chkTai.AutoSize = true;
            this.chkTai.Location = new System.Drawing.Point(89, 28);
            this.chkTai.Name = "chkTai";
            this.chkTai.Size = new System.Drawing.Size(43, 23);
            this.chkTai.TabIndex = 1;
            this.chkTai.Text = "退";
            this.chkTai.UseVisualStyleBackColor = true;
            // 
            // chkGen
            // 
            this.chkGen.AutoSize = true;
            this.chkGen.Location = new System.Drawing.Point(24, 28);
            this.chkGen.Name = "chkGen";
            this.chkGen.Size = new System.Drawing.Size(43, 23);
            this.chkGen.TabIndex = 0;
            this.chkGen.Text = "現";
            this.chkGen.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(12, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(350, 46);
            this.button1.TabIndex = 2;
            this.button1.Text = "実行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(12, 166);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(350, 46);
            this.button2.TabIndex = 3;
            this.button2.Text = "戻る";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 235);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Meiryo UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "組合員エクセル予定申告書発行";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkHi;
        private System.Windows.Forms.CheckBox chkFumei;
        private System.Windows.Forms.CheckBox chkKyu;
        private System.Windows.Forms.CheckBox chkTai;
        private System.Windows.Forms.CheckBox chkGen;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

