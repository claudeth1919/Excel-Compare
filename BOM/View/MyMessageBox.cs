using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BOM.Tool;
using System.IO;

namespace BOM.View
{
    public partial class MyMessageBox : Form
    {
        public AlarmType Type;
        public bool IsClosed = false;
        public System.Windows.Forms.RichTextBox MyRichTextBox;
        public MyMessageBox(AlarmType type, List<string> messages)
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            Type = type;
            string currentPath = Path.GetDirectoryName(Application.ExecutablePath);
            switch (type)
            {
                case AlarmType.ERROR:
                    Picture.Image = Image.FromFile($@"{currentPath}\resourse\error.png");
                    MessBox.Font = new System.Drawing.Font("Arial Narrow", 13, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    break;
                case AlarmType.SUCCESS:
                    Picture.Image = Image.FromFile($@"{currentPath}\resourse\success.png");
                    MessBox.Font = new System.Drawing.Font("Arial Narrow", 13, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    break;
                case AlarmType.LOADING:
                    Picture.Image = Image.FromFile($@"{currentPath}\resourse\clock.png");
                    MessBox.Font = new System.Drawing.Font("Arial Narrow", 13, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    break;
                case AlarmType.WARNING:
                    Picture.Image = Image.FromFile($@"{currentPath}\resourse\warning.png");
                    MessBox.Font = new System.Drawing.Font("Arial Narrow", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    break;
                default:
                    break;
            }
            AddText(messages);
            MyRichTextBox = MessBox;
        }

        public void AddText(List<string> messages)
        {
            MessBox.ReadOnly = true;
            foreach (string message in messages)
            {
                MessBox.AppendText(message+"\n");
            }
        }

        private void Click_Ok(object sender, EventArgs e)
        {
            this.Close();
            IsClosed = true;
        }

        private void Closed_Event(object sender, FormClosedEventArgs e)
        {
            IsClosed = true;
        }
    }
}
