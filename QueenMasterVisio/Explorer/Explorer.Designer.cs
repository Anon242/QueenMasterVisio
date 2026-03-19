namespace QueenMasterVisio
{
    partial class Explorer
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
            this.components = new System.ComponentModel.Container();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.переименоватьF2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.удалитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBox1.Location = new System.Drawing.Point(0, 0);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(493, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.Click += new System.EventHandler(this.textBox1_Click);
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            this.textBox1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyUp);
            // 
            // listView1
            // 
            this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listView1.Font = new System.Drawing.Font("Cascadia Mono", 8.25F);
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(0, 20);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(493, 430);
            this.listView1.TabIndex = 2;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.DrawItem += new System.Windows.Forms.DrawListViewItemEventHandler(this.listView1_DrawItem);
            this.listView1.Click += new System.EventHandler(this.listView1_Click);
            this.listView1.DoubleClick += new System.EventHandler(this.listView1_DoubleClick);
            this.listView1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.listView1_KeyUp);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.переименоватьF2ToolStripMenuItem,
            this.toolStripSeparator1,
            this.удалитьToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(181, 54);
            // 
            // переименоватьF2ToolStripMenuItem
            // 
            this.переименоватьF2ToolStripMenuItem.Name = "переименоватьF2ToolStripMenuItem";
            this.переименоватьF2ToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F2;
            this.переименоватьF2ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.переименоватьF2ToolStripMenuItem.Text = "Переименовать";
            this.переименоватьF2ToolStripMenuItem.Click += new System.EventHandler(this.переименоватьF2ToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(177, 6);
            // 
            // удалитьToolStripMenuItem
            // 
            this.удалитьToolStripMenuItem.Name = "удалитьToolStripMenuItem";
            this.удалитьToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.удалитьToolStripMenuItem.Text = "Удалить";
            this.удалитьToolStripMenuItem.Click += new System.EventHandler(this.удалитьToolStripMenuItem_Click);
            // 
            // Explorer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.textBox1);
            this.Name = "Explorer";
            this.Size = new System.Drawing.Size(493, 450);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem переименоватьF2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem удалитьToolStripMenuItem;
    }
}