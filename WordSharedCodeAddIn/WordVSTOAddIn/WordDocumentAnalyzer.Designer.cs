namespace WordVSTOAddIn
{
    partial class WordDocumentAnalyzer
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnShowText = new System.Windows.Forms.Button();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnShowText
            // 
            this.btnShowText.Location = new System.Drawing.Point(61, 32);
            this.btnShowText.Name = "btnShowText";
            this.btnShowText.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnShowText.Size = new System.Drawing.Size(200, 23);
            this.btnShowText.TabIndex = 0;
            this.btnShowText.Text = "Show Text";
            this.btnShowText.UseVisualStyleBackColor = true;
            this.btnShowText.Click += new System.EventHandler(this.btnShowText_Click);
            // 
            // txtResult
            // 
            this.txtResult.Location = new System.Drawing.Point(42, 102);
            this.txtResult.Multiline = true;
            this.txtResult.Name = "txtResult";
            this.txtResult.ReadOnly = true;
            this.txtResult.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtResult.Size = new System.Drawing.Size(226, 102);
            this.txtResult.TabIndex = 1;
            // 
            // WordDocumentAnalyzer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.btnShowText);
            this.Name = "WordDocumentAnalyzer";
            this.Size = new System.Drawing.Size(294, 234);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnShowText;
        private System.Windows.Forms.TextBox txtResult;
    }
}
