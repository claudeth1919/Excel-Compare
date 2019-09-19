namespace BOM.View
{
    partial class MyMessageBox
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
            this.OkBtn = new System.Windows.Forms.Button();
            this.MessBox = new System.Windows.Forms.RichTextBox();
            this.Picture = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.Picture)).BeginInit();
            this.SuspendLayout();
            // 
            // OkBtn
            // 
            this.OkBtn.Location = new System.Drawing.Point(229, 178);
            this.OkBtn.Name = "OkBtn";
            this.OkBtn.Size = new System.Drawing.Size(75, 23);
            this.OkBtn.TabIndex = 0;
            this.OkBtn.Text = "Ok";
            this.OkBtn.UseVisualStyleBackColor = true;
            this.OkBtn.Click += new System.EventHandler(this.Click_Ok);
            // 
            // MessBox
            // 
            this.MessBox.Location = new System.Drawing.Point(148, 25);
            this.MessBox.Name = "MessBox";
            this.MessBox.Size = new System.Drawing.Size(351, 147);
            this.MessBox.TabIndex = 3;
            this.MessBox.Text = "";
            // 
            // Picture
            // 
            this.Picture.Location = new System.Drawing.Point(12, 29);
            this.Picture.Name = "Picture";
            this.Picture.Size = new System.Drawing.Size(130, 147);
            this.Picture.TabIndex = 4;
            this.Picture.TabStop = false;
            // 
            // MyMessageBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 213);
            this.Controls.Add(this.Picture);
            this.Controls.Add(this.MessBox);
            this.Controls.Add(this.OkBtn);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MyMessageBox";
            this.Text = "Mensaje";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Closed_Event);
            ((System.ComponentModel.ISupportInitialize)(this.Picture)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button OkBtn;
        private System.Windows.Forms.RichTextBox MessBox;
        private System.Windows.Forms.PictureBox Picture;
    }
}