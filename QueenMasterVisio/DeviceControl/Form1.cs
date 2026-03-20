using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;

namespace QueenMasterVisio
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
            SetupDataGridView();
            LoadData();
        }
        private void LoadData()
        {
            // Пример данных: список объектов
            var items = new[]
            {
                new MyItem { Code = 101, Description = "Красный квадрат", Shape = "Квадрат", Page = 12 },
                new MyItem { Code = 102, Description = "Синий кругкругкругкруг\nкругкруг", Shape = "Круг", Page = 15 },
                new MyItem { Code = 103, Description = "Зелёный треугольник", Shape = "Треугольник", Page = 20 }
            };

            for(int i = 0; i < 20; i++)
            foreach (var item in items)
            {
                int rowIndex = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[rowIndex];

                // Заполняем столбцы
                row.Cells["IconColumn"].Value = item.Icon;          // значок
                row.Cells["CodeColumn"].Value = item.Code;          // код (скрытый)
                row.Cells["DescriptionColumn"].Value = item.Description;
                row.Cells["ShapeColumn"].Value = item.Shape;
                row.Cells["PageColumn"].Value = item.Page;
            }

            // Альтернативный способ: привязка к DataTable (см. комментарий ниже)
        }
        private void SetupDataGridView()
        {
            // Столбец для значка
            DataGridViewImageColumn iconColumn = new DataGridViewImageColumn();
            iconColumn.Name = "IconColumn";
            iconColumn.HeaderText = "";
            iconColumn.ImageLayout = DataGridViewImageCellLayout.NotSet;
            iconColumn.Width = 30;
            dataGridView1.Columns.Add(iconColumn);

            // Столбец "Код" (скрытый)
            DataGridViewTextBoxColumn codeColumn = new DataGridViewTextBoxColumn();
            codeColumn.Name = "CodeColumn";
            codeColumn.HeaderText = "Код";
            codeColumn.Visible = false;
            dataGridView1.Columns.Add(codeColumn);

            // Столбец "Описание"
            DataGridViewTextBoxColumn descriptionColumn = new DataGridViewTextBoxColumn();
            descriptionColumn.Name = "DescriptionColumn";
            descriptionColumn.HeaderText = "Описание";
            descriptionColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            descriptionColumn.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns.Add(descriptionColumn);

            // Столбец "Фигура"
            DataGridViewTextBoxColumn shapeColumn = new DataGridViewTextBoxColumn();
            shapeColumn.Name = "ShapeColumn";
            shapeColumn.HeaderText = "Фигура";
            shapeColumn.Width = 100;
            dataGridView1.Columns.Add(shapeColumn);

            // Столбец "Страница"
            DataGridViewTextBoxColumn pageColumn = new DataGridViewTextBoxColumn();
            pageColumn.Name = "PageColumn";
            pageColumn.HeaderText = "Стр";
            dataGridView1.Columns.Add(pageColumn);
            pageColumn.Width = 40;

        }
    }

    // Вспомогательный класс для хранения данных
    public class MyItem
    {
        public int Code { get; set; }
        public string Description { get; set; }
        public string Shape { get; set; }
        public int Page { get; set; }
        public Image Icon { get; set; }
    }
}