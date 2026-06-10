using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace QueenMasterVisio.ChangeLog
{
    public partial class ChangeLog : UserControl
    {
        Microsoft.Office.Interop.Visio.Application app;

        private string changeLogPath = "";
        private Microsoft.Office.Interop.Visio.Window myWindow;
        private DateTime date;
        public ChangeLog(Microsoft.Office.Interop.Visio.Window window,Microsoft.Office.Interop.Visio.Application app, string changeLogPath)
        {
            InitializeComponent();
            this.myWindow = window;
            this.changeLogPath = changeLogPath;
            this.app = app;
            this.date = DateTime.Now;
        }


        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            e.Graphics.DrawString("Ваш текст", this.Font, Brushes.Black, 10, 10);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                textBox1.Text = "null";
            }

            ChangeModel model = new ChangeModel();
            model.name = textBox1.Text;
            model.description = richTextBox1.Text;
            model.date = date;
            model.author = app.UserName;

            string fileName = "Changelog_" + model.author +"_" + model.date.ToString("yyyyMMdd_HHmm");
            string text = $"{model.name}\n{model.description}\n{model.version}";

            SendChangeLog(fileName,text);
            myWindow.Close();
        }

        private void SendChangeLog(string fileName, string text)
        {
            try
            {
                System.IO.File.WriteAllText(System.IO.Path.Combine(changeLogPath, fileName) , text);
                System.IO.File.SetAttributes(System.IO.Path.Combine(changeLogPath, fileName), System.IO.File.GetAttributes(System.IO.Path.Combine(changeLogPath, fileName)) | System.IO.FileAttributes.ReadOnly);
            }
            catch (Exception)
            {

              
            }
            
        }

    }
}
