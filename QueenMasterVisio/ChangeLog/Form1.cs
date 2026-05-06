using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueenMasterVisio.ChangeLog
{
    public partial class Form1 : Form
    {
        private bool closeFlag = false;
        private string changeLogPath = "";

        public Form1(string changeLogPath)
        {
            InitializeComponent();
            this.changeLogPath = changeLogPath;
        }
    

        public void ChangeLog(string text)
        {
            label3.Text = text;
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            e.Graphics.DrawString("Ваш текст", this.Font, Brushes.Black, 10, 10);
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(!closeFlag)
                SendChangeLog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendChangeLog();
            closeFlag = true;
            this.Close();
        }

        private void SendChangeLog()
        {
            string supertext = "2026-13-1\nЯ сделал тото тото\nВот пруфы лог лог лог\nлог лог лог";
            //System.IO.File.WriteAllText(changeLogPath, supertext);
        }
    }
}
