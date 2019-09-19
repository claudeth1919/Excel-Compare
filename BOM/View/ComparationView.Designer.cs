namespace BOM.View
{
    partial class ComparationView
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
            this.Tittle = new System.Windows.Forms.Label();
            this.Panel = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.FileNameLabel_1 = new System.Windows.Forms.Label();
            this.FileNameLabel_2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Tittle
            // 
            this.Tittle.AutoSize = true;
            this.Tittle.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F);
            this.Tittle.Location = new System.Drawing.Point(53, 38);
            this.Tittle.Name = "Tittle";
            this.Tittle.Size = new System.Drawing.Size(151, 31);
            this.Tittle.TabIndex = 0;
            this.Tittle.Text = "Diferencias";
            // 
            // Panel
            // 
            this.Panel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.Panel.Location = new System.Drawing.Point(31, 138);
            this.Panel.Name = "Panel";
            this.Panel.Size = new System.Drawing.Size(863, 510);
            this.Panel.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(56, 79);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "File 1:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(56, 106);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "File 2:";
            // 
            // FileNameLabel_1
            // 
            this.FileNameLabel_1.AutoSize = true;
            this.FileNameLabel_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.FileNameLabel_1.Location = new System.Drawing.Point(122, 79);
            this.FileNameLabel_1.Name = "FileNameLabel_1";
            this.FileNameLabel_1.Size = new System.Drawing.Size(38, 18);
            this.FileNameLabel_1.TabIndex = 4;
            this.FileNameLabel_1.Text = "file 1";
            // 
            // FileNameLabel_2
            // 
            this.FileNameLabel_2.AutoSize = true;
            this.FileNameLabel_2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.FileNameLabel_2.Location = new System.Drawing.Point(122, 106);
            this.FileNameLabel_2.Name = "FileNameLabel_2";
            this.FileNameLabel_2.Size = new System.Drawing.Size(38, 18);
            this.FileNameLabel_2.TabIndex = 5;
            this.FileNameLabel_2.Text = "file 2";
            // 
            // ComparationView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1089, 679);
            this.Controls.Add(this.FileNameLabel_2);
            this.Controls.Add(this.FileNameLabel_1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Panel);
            this.Controls.Add(this.Tittle);
            this.Name = "ComparationView";
            this.Text = "AssignOrder";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Tittle;
        private System.Windows.Forms.FlowLayoutPanel Panel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label FileNameLabel_1;
        private System.Windows.Forms.Label FileNameLabel_2;
    }
}