namespace ShadowEdit
{
    partial class FormError
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormError));
            this.TextBoxError = new System.Windows.Forms.TextBox();
            this.ButtonOK = new System.Windows.Forms.Button();
            this.ButtonCopyToClipboard = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TextBoxError
            // 
            this.TextBoxError.Location = new System.Drawing.Point(12, 12);
            this.TextBoxError.Multiline = true;
            this.TextBoxError.Name = "TextBoxError";
            this.TextBoxError.ReadOnly = true;
            this.TextBoxError.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.TextBoxError.Size = new System.Drawing.Size(375, 311);
            this.TextBoxError.TabIndex = 0;
            this.TextBoxError.WordWrap = false;
            // 
            // ButtonOK
            // 
            this.ButtonOK.Location = new System.Drawing.Point(312, 333);
            this.ButtonOK.Name = "ButtonOK";
            this.ButtonOK.Size = new System.Drawing.Size(75, 23);
            this.ButtonOK.TabIndex = 1;
            this.ButtonOK.Text = "OK";
            this.ButtonOK.UseVisualStyleBackColor = true;
            this.ButtonOK.Click += new System.EventHandler(this.ButtonOK_Click);
            // 
            // ButtonCopyToClipboard
            // 
            this.ButtonCopyToClipboard.Location = new System.Drawing.Point(189, 333);
            this.ButtonCopyToClipboard.Name = "ButtonCopyToClipboard";
            this.ButtonCopyToClipboard.Size = new System.Drawing.Size(117, 23);
            this.ButtonCopyToClipboard.TabIndex = 2;
            this.ButtonCopyToClipboard.Text = "Copy To Clipboard";
            this.ButtonCopyToClipboard.UseVisualStyleBackColor = true;
            this.ButtonCopyToClipboard.Click += new System.EventHandler(this.ButtonCopyToClipboard_Click);
            // 
            // FormError
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 368);
            this.Controls.Add(this.ButtonCopyToClipboard);
            this.Controls.Add(this.ButtonOK);
            this.Controls.Add(this.TextBoxError);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormError";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Error!";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ButtonOK;
        private System.Windows.Forms.Button ButtonCopyToClipboard;
        public System.Windows.Forms.TextBox TextBoxError;
    }
}