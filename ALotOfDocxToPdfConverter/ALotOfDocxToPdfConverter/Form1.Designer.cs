namespace WordToPdf
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.SelectMultFileButton = new System.Windows.Forms.Button();
            this.processLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.OpenAfterExport = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // SelectMultFileButton
            // 
            this.SelectMultFileButton.Font = new System.Drawing.Font("新細明體", 12F);
            this.SelectMultFileButton.Location = new System.Drawing.Point(15, 139);
            this.SelectMultFileButton.Name = "SelectMultFileButton";
            this.SelectMultFileButton.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.SelectMultFileButton.Size = new System.Drawing.Size(732, 87);
            this.SelectMultFileButton.TabIndex = 0;
            this.SelectMultFileButton.Text = "Select Mult .docx File";
            this.SelectMultFileButton.UseVisualStyleBackColor = true;
            this.SelectMultFileButton.Click += new System.EventHandler(this.SelectMultFileButton_Click);
            // 
            // processLabel
            // 
            this.processLabel.Font = new System.Drawing.Font("新細明體", 12F);
            this.processLabel.Location = new System.Drawing.Point(12, 9);
            this.processLabel.Name = "processLabel";
            this.processLabel.Size = new System.Drawing.Size(735, 98);
            this.processLabel.TabIndex = 1;
            this.processLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 110);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(735, 23);
            this.progressBar.TabIndex = 2;
            // 
            // OpenAfterExport
            // 
            this.OpenAfterExport.AutoSize = true;
            this.OpenAfterExport.Font = new System.Drawing.Font("新細明體", 12F);
            this.OpenAfterExport.Location = new System.Drawing.Point(12, 235);
            this.OpenAfterExport.Name = "OpenAfterExport";
            this.OpenAfterExport.Size = new System.Drawing.Size(135, 20);
            this.OpenAfterExport.TabIndex = 3;
            this.OpenAfterExport.Text = "OpenAfterExport";
            this.OpenAfterExport.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(759, 286);
            this.Controls.Add(this.OpenAfterExport);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.processLabel);
            this.Controls.Add(this.SelectMultFileButton);
            this.Name = "Form1";
            this.Text = "A Lot Of Docx To Pdf Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SelectMultFileButton;
        private System.Windows.Forms.Label processLabel;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.CheckBox OpenAfterExport;
    }
}

