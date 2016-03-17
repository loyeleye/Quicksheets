namespace DoMyTimesheet
{
    partial class ShowEntry
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
            this.title = new System.Windows.Forms.Label();
            this.textBox = new System.Windows.Forms.RichTextBox();
            this.okButton = new System.Windows.Forms.Button();
            this.applyButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // title
            // 
            this.title.AutoSize = true;
            this.title.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.title.Font = new System.Drawing.Font("Sitka Heading", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.title.Location = new System.Drawing.Point(12, 13);
            this.title.Name = "title";
            this.title.Size = new System.Drawing.Size(109, 30);
            this.title.TabIndex = 0;
            this.title.Text = "Timesheet";
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(13, 47);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(670, 53);
            this.textBox.TabIndex = 1;
            this.textBox.Text = "";
            // 
            // okButton
            // 
            this.okButton.BackColor = System.Drawing.SystemColors.Window;
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(17, 106);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 2;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = false;
            // 
            // applyButton
            // 
            this.applyButton.BackColor = System.Drawing.SystemColors.Window;
            this.applyButton.DialogResult = System.Windows.Forms.DialogResult.Retry;
            this.applyButton.Location = new System.Drawing.Point(608, 106);
            this.applyButton.Name = "applyButton";
            this.applyButton.Size = new System.Drawing.Size(75, 23);
            this.applyButton.TabIndex = 3;
            this.applyButton.Text = "Apply";
            this.applyButton.UseVisualStyleBackColor = false;
            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.SystemColors.Window;
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(98, 105);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = false;
            // 
            // ShowEntry
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(693, 140);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.applyButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.title);
            this.Name = "ShowEntry";
            this.Text = "Show Entry";
            this.TransparencyKey = System.Drawing.Color.Red;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public string TextField
        {
            get
            {
                return this.textBox.Text.Trim();
            }
            set
            {
                this.textBox.Text = value;
            }
        }

        public string TitleLabel
        {
            set
            {
                this.title.Text = value.Trim().Replace("\t", "   "); 
            }
        }

        private System.Windows.Forms.Label title;
        private System.Windows.Forms.RichTextBox textBox;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button applyButton;
        private System.Windows.Forms.Button cancelButton;
    }
}