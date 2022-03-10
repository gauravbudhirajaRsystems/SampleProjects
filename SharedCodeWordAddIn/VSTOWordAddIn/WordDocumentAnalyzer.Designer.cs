namespace VSTOWordAddIn
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
            this.tblLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.btnUnicode = new System.Windows.Forms.Button();
            this.btnCharCount = new System.Windows.Forms.Button();
            this.btnWordCount = new System.Windows.Forms.Button();
            this.txtUnicode = new System.Windows.Forms.TextBox();
            this.txtCharCount = new System.Windows.Forms.TextBox();
            this.txtWordCount = new System.Windows.Forms.TextBox();
            this.tblLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // tblLayoutPanel
            // 
            this.tblLayoutPanel.ColumnCount = 2;
            this.tblLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 37.87879F));
            this.tblLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 62.12121F));
            this.tblLayoutPanel.Controls.Add(this.txtWordCount, 1, 2);
            this.tblLayoutPanel.Controls.Add(this.txtCharCount, 1, 1);
            this.tblLayoutPanel.Controls.Add(this.btnUnicode, 0, 0);
            this.tblLayoutPanel.Controls.Add(this.btnCharCount, 0, 1);
            this.tblLayoutPanel.Controls.Add(this.btnWordCount, 0, 2);
            this.tblLayoutPanel.Controls.Add(this.txtUnicode, 1, 0);
            this.tblLayoutPanel.Location = new System.Drawing.Point(3, 3);
            this.tblLayoutPanel.Name = "tblLayoutPanel";
            this.tblLayoutPanel.RowCount = 3;
            this.tblLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.7014F));
            this.tblLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.2986F));
            this.tblLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 278F));
            this.tblLayoutPanel.Size = new System.Drawing.Size(594, 778);
            this.tblLayoutPanel.TabIndex = 0;
            // 
            // btnUnicode
            // 
            this.btnUnicode.AllowDrop = true;
            this.btnUnicode.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnUnicode.Location = new System.Drawing.Point(25, 91);
            this.btnUnicode.Name = "btnUnicode";
            this.btnUnicode.Size = new System.Drawing.Size(174, 71);
            this.btnUnicode.TabIndex = 0;
            this.btnUnicode.Text = "Show Unicode";
            this.btnUnicode.UseVisualStyleBackColor = true;
            this.btnUnicode.Click += new System.EventHandler(this.btnUnicode_Click);
            // 
            // btnCharCount
            // 
            this.btnCharCount.AllowDrop = true;
            this.btnCharCount.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnCharCount.Location = new System.Drawing.Point(25, 340);
            this.btnCharCount.Name = "btnCharCount";
            this.btnCharCount.Size = new System.Drawing.Size(174, 71);
            this.btnCharCount.TabIndex = 1;
            this.btnCharCount.Text = "Show Char Count";
            this.btnCharCount.UseVisualStyleBackColor = true;
            this.btnCharCount.Click += new System.EventHandler(this.btnCharCount_Click);
            // 
            // btnWordCount
            // 
            this.btnWordCount.AllowDrop = true;
            this.btnWordCount.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnWordCount.Location = new System.Drawing.Point(25, 603);
            this.btnWordCount.Name = "btnWordCount";
            this.btnWordCount.Size = new System.Drawing.Size(174, 71);
            this.btnWordCount.TabIndex = 2;
            this.btnWordCount.Text = "Show Word Count";
            this.btnWordCount.UseVisualStyleBackColor = true;
            this.btnWordCount.Click += new System.EventHandler(this.btnWordCount_Click);
            // 
            // txtUnicode
            // 
            this.txtUnicode.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtUnicode.Enabled = false;
            this.txtUnicode.Location = new System.Drawing.Point(236, 24);
            this.txtUnicode.Multiline = true;
            this.txtUnicode.Name = "txtUnicode";
            this.txtUnicode.ReadOnly = true;
            this.txtUnicode.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtUnicode.Size = new System.Drawing.Size(346, 205);
            this.txtUnicode.TabIndex = 3;
            // 
            // txtCharCount
            // 
            this.txtCharCount.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtCharCount.Enabled = false;
            this.txtCharCount.Location = new System.Drawing.Point(236, 273);
            this.txtCharCount.Multiline = true;
            this.txtCharCount.Name = "txtCharCount";
            this.txtCharCount.ReadOnly = true;
            this.txtCharCount.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtCharCount.Size = new System.Drawing.Size(346, 205);
            this.txtCharCount.TabIndex = 4;
            // 
            // txtWordCount
            // 
            this.txtWordCount.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtWordCount.Enabled = false;
            this.txtWordCount.Location = new System.Drawing.Point(236, 536);
            this.txtWordCount.Multiline = true;
            this.txtWordCount.Name = "txtWordCount";
            this.txtWordCount.ReadOnly = true;
            this.txtWordCount.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtWordCount.Size = new System.Drawing.Size(346, 205);
            this.txtWordCount.TabIndex = 5;
            // 
            // WordDocumentAnalyzer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tblLayoutPanel);
            this.Name = "WordDocumentAnalyzer";
            this.Size = new System.Drawing.Size(600, 760);
            this.tblLayoutPanel.ResumeLayout(false);
            this.tblLayoutPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tblLayoutPanel;
        private System.Windows.Forms.Button btnUnicode;
        private System.Windows.Forms.Button btnCharCount;
        private System.Windows.Forms.Button btnWordCount;
        private System.Windows.Forms.TextBox txtUnicode;
        private System.Windows.Forms.TextBox txtWordCount;
        private System.Windows.Forms.TextBox txtCharCount;
    }
}
