using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueenMasterVisio.ChangeLog
{
    public partial class ChangeLog : UserControl
    {

        private string changeLogPath = "";
        private Microsoft.Office.Interop.Visio.Window myWindow;
        public ChangeLog(Microsoft.Office.Interop.Visio.Window window, string changeLogPath)
        {
            InitializeComponent();
            this.myWindow = window;
            this.changeLogPath = changeLogPath;
        }


        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            e.Graphics.DrawString("Ваш текст", this.Font, Brushes.Black, 10, 10);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendChangeLog();
            myWindow.Close();
        }

        private void SendChangeLog()
        {
            try
            {
                string supertext = "2026-13-1\nЯ сделал тото тото\nВот пруфы лог лог лог\nлог лог лог";
                System.IO.File.WriteAllText(changeLogPath, supertext);
            }
            catch (Exception)
            {

              
            }
            
        }

    }
}
