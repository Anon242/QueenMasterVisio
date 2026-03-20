using System.Windows.Forms;

namespace QueenMasterVisio
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // Объявляем все элементы управления, которые будут созданы в дизайнере
        private System.Windows.Forms.TableLayoutPanel topPanel;
        private System.Windows.Forms.ComboBox cmbScope;
        private System.Windows.Forms.Button btnErrors;
        private System.Windows.Forms.Button btnWarnings;
        private System.Windows.Forms.Button btnMessages;
        private System.Windows.Forms.Button btnCheckPage;
        private System.Windows.Forms.DataGridView dataGridView1;

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
            this.topPanel = new System.Windows.Forms.TableLayoutPanel();
            this.cmbScope = new System.Windows.Forms.ComboBox();
            this.btnErrors = new System.Windows.Forms.Button();
            this.btnWarnings = new System.Windows.Forms.Button();
            this.btnMessages = new System.Windows.Forms.Button();
            this.btnCheckPage = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.topPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // topPanel
            // 
            this.topPanel.ColumnCount = 5;
            this.topPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.topPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.topPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.topPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.topPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.topPanel.Controls.Add(this.cmbScope, 0, 0);
            this.topPanel.Controls.Add(this.btnErrors, 1, 0);
            this.topPanel.Controls.Add(this.btnWarnings, 2, 0);
            this.topPanel.Controls.Add(this.btnMessages, 3, 0);
            this.topPanel.Controls.Add(this.btnCheckPage, 4, 0);
            this.topPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.topPanel.Location = new System.Drawing.Point(0, 0);
            this.topPanel.Name = "topPanel";
            this.topPanel.RowCount = 1;
            this.topPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.topPanel.Size = new System.Drawing.Size(800, 40);
            this.topPanel.TabIndex = 0;
            // 
            // cmbScope
            // 
            this.cmbScope.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbScope.FormattingEnabled = true;
            this.cmbScope.Items.AddRange(new object[] {
            "Документ",
            "Страница"});
            this.cmbScope.Location = new System.Drawing.Point(3, 3);
            this.cmbScope.Name = "cmbScope";
            this.cmbScope.Size = new System.Drawing.Size(154, 21);
            this.cmbScope.TabIndex = 0;
            // 
            // btnErrors
            // 
            this.btnErrors.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnErrors.Location = new System.Drawing.Point(163, 3);
            this.btnErrors.Name = "btnErrors";
            this.btnErrors.Size = new System.Drawing.Size(154, 34);
            this.btnErrors.TabIndex = 1;
            this.btnErrors.Text = "0 Ошибка";
            this.btnErrors.UseVisualStyleBackColor = true;
            // 
            // btnWarnings
            // 
            this.btnWarnings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnWarnings.Location = new System.Drawing.Point(323, 3);
            this.btnWarnings.Name = "btnWarnings";
            this.btnWarnings.Size = new System.Drawing.Size(154, 34);
            this.btnWarnings.TabIndex = 2;
            this.btnWarnings.Text = "0 Предупреждения";
            this.btnWarnings.UseVisualStyleBackColor = true;
            // 
            // btnMessages
            // 
            this.btnMessages.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnMessages.Location = new System.Drawing.Point(483, 3);
            this.btnMessages.Name = "btnMessages";
            this.btnMessages.Size = new System.Drawing.Size(154, 34);
            this.btnMessages.TabIndex = 3;
            this.btnMessages.Text = "0 Сообщения";
            this.btnMessages.UseVisualStyleBackColor = true;
            // 
            // btnCheckPage
            // 
            this.btnCheckPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnCheckPage.Location = new System.Drawing.Point(643, 3);
            this.btnCheckPage.Name = "btnCheckPage";
            this.btnCheckPage.Size = new System.Drawing.Size(154, 34);
            this.btnCheckPage.TabIndex = 4;
            this.btnCheckPage.Text = "Проверить страницу";
            this.btnCheckPage.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 40);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(800, 410);
            this.dataGridView1.TabIndex = 1;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.topPanel);
            this.Name = "Form1";
            this.Text = "QueenMasterVisio";
            this.topPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
    }
}