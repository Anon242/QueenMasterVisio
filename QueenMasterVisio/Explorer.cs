using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Reflection.Emit;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Color = System.Drawing.Color;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio
{
    public partial class Explorer : UserControl
    {
        private Visio.Application visioApp;
        private List<PageElement> pageList = new List<PageElement>();
        List<string> titleList = new List<string>() {"ИНФОРМАЦИЯ", "ПЛАН ТРАССИРОВОК","DEVICES","LIGHTS","BACKGROUNDS"};
        int titleListNum = 0;
        private Visio.Window customWindow;

        // Переменные для поисковика
        private bool isFound = false;
        private int index = 0;
        List<ListViewItem> publicListViewItemsForEnterKey = new List<ListViewItem>();

        List<Color> gradient = new List<Color>() 
        {
            Color.FromArgb(255, 255,127,0),
            Color.FromArgb(255, 227,26,28),
            Color.FromArgb(255,51,160,44),
            Color.FromArgb(255,31,120,180),
        };

        public Explorer(Visio.Application application, Visio.Window _customWindow)
        {
            customWindow = _customWindow;
            visioApp = application;
            InitializeComponent();

            listView1.FullRowSelect = true;
            listView1.Columns.Add("Стрaница", 34);
            listView1.Columns.Add("Название страницы", 400);
            listView1.Columns.Add("Тип", 50);

            // Отключаем горизонтальную прокрутку
            listView1.Scrollable = true;


            listView1.AllowDrop = true;
            listView1.ItemDrag += ListView1_ItemDrag;
            listView1.DragEnter += ListView1_DragEnter;
            listView1.DragOver += ListView1_DragOver;
            listView1.DragDrop += ListView1_DragDrop;

            // Включаем перетаскивание
            listView1.AllowDrop = true;
            listView1.ContextMenuStrip = contextMenuStrip1;

        }
        private void ListView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            // Начало перетаскивания
            if (e.Button == MouseButtons.Left)
            {
                listView1.DoDragDrop(e.Item, DragDropEffects.Move);
            }
        }

        private void ListView1_DragEnter(object sender, DragEventArgs e)
        {
            // Разрешаем перетаскивание
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void ListView1_DragOver(object sender, DragEventArgs e)
        {
            // Показываем где можно бросить
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void ListView1_DragDrop(object sender, DragEventArgs e)
        {
            // Обработка "бросания" элемента
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                ListViewItem draggedItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));
                Point clientPoint = listView1.PointToClient(new Point(e.X, e.Y));
                ListViewItem targetItem = listView1.GetItemAt(clientPoint.X, clientPoint.Y);

                if (targetItem != null && draggedItem != targetItem)
                {
                    Visio.Page draggedPage = visioApp.ActiveDocument.Pages.ItemU[draggedItem.SubItems[2].Text];
                    Visio.Page targePage = visioApp.ActiveDocument.Pages.ItemU[targetItem.SubItems[2].Text];

                    DialogResult result = MessageBox.Show(
    $"Переместить страницу '{draggedItem.SubItems[1].Text}' на страницу '{targePage.Index - 1}'?",
    "Подтверждение перемещения",
    MessageBoxButtons.OKCancel,
    MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.OK)
                    {
                        draggedPage.Index = (short)(targePage.Index - 1);
                    }
                    //MoveListViewItem(draggedItem, targetItem);
                    UpdateExplorer(visioApp.ActiveDocument);

                }
            }
        }


        public void LookDevices(Visio.Page page)
        {
            List<string> shapes = new List<string>();
            int shapeCount = 0;
            int pageCount = 0;
            foreach (Shape shape in page.Shapes)
            {
                // Мда блять кодер конечно нашелся
                if(shape.Name.Contains("Device"))
                {
                    shapes.Add('G' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + " ");
                    shapes.Add('G' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + ",");
                    shapes.Add('G' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + "-");
                    shapeCount++;
                }

                else if (shape.Name.Contains("Light"))
                {
                    shapes.Add('L' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + " ");
                    shapes.Add('L' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + ",");
                    shapes.Add('L' + Tools.CellFormulaGet(shape, "Prop.Number").Replace("\"", "") + "-");
                    shapeCount++;
                }

            }

            if (shapes.Count == 0)
                return;

            textBox1.Text = "";
            HighlightItemsColorReset(listView1);

            bool one = false;
            foreach (string shape in shapes)
            {
                Debug.WriteLine(shape);
                List<ListViewItem> result = SearchItem(shape, listView1);
                if (result.Count == 0)
                    continue;

                pageCount++;
                HighlightItemsInList(result, listView1, Color.LightBlue, Color.Black);
                if (!one)
                {
                    listView1.TopItem = result[0];
                    one = true;
                }
            }
            customWindow.Caption = "Найдено: "+shapeCount + " устр. на "+ pageCount + " вх.";

        }

        public void AddElement(string name)
        {
            //listBox1.Invoke(new Action(() =>  listBox1.Items.Add(name)));
        }

        // Мы обновляем список в 4 случаях
        // При создании страницы
        // При удалении страницы
        // Документ опен
        // И раз в 5 сек проверяем все страницы

        public void UpdateExplorer(Document doc)
        {
            Thread th = new Thread(() =>
            {
                
                UpdateExplorerElements();
               
            });
            th.IsBackground = true;
            th.Start();
        }

        private void GoToPageByNameU(string nameU)
        {
            try
            {
                Debug.WriteLine(nameU);
                if (!titleList.Any(title => nameU.Contains(title)))
                {
                    Debug.WriteLine(nameU);

                    Visio.Page targetPage = visioApp.ActiveDocument.Pages.ItemU[nameU];
                    if (visioApp.ActiveWindow.Page != targetPage)
                    {
                        visioApp.ActiveWindow.Page = targetPage;
                        //visioApp.ActiveWindow.Page.Application.ActiveWindow.SetViewRect(0, 12, 13, 13);
                    }
                }
            }
            catch
            {
                Debug.WriteLine("Не удалось перейти на страницу: " + nameU);
            }
        }
        private Bitmap CreateColorIcon(System.Drawing.Color color)
        {
            var bmp = new Bitmap(2, 12);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(color);
            }
            return bmp;
        }
        private void UpdateExplorerElements()
        {
            int topIndex = 0;
            bool buffer = false;
            // Запоминаем
            if (listView1.Items.Count != 0)
            {
                buffer = true;
                listView1.Invoke(new Action(() =>
                {
                    topIndex = listView1.TopItem.Index;
                }));
    
            }

            listView1.Invoke(new Action(() => listView1.Items.Clear()));

            Array namesArray;
            Array namesArrayU;

            visioApp.ActiveDocument.Pages.GetNames(out namesArray);
            visioApp.ActiveDocument.Pages.GetNamesU(out namesArrayU);

            int index = 0;

            listView1.Invoke(new Action(() =>
            {
                listView1.SmallImageList = new ImageList();
                listView1.SmallImageList.ImageSize = new Size(2, 12);
            }));

            string regexPlanBufferString = "";
            int regexPlanIndexColor = 0;

            foreach (string page in namesArray)
                    listView1.Invoke(new Action(() => 
                    {
                        index++;
                        var item = new ListViewItem(index.ToString());

                        int oldTitleNum = titleListNum;
                        Regex regexDevice = new Regex(@"^G\d");
                        Regex regexLight = new Regex(@"^L\d");
                        if (regexDevice.IsMatch(page))
                        {
                            titleListNum = 2;
                        }
                        else if (regexLight.IsMatch(page))
                        {
                            titleListNum = 3;
                        }
                        else if (page == "фон" || page.Contains("DEVELOP "))
                        {
                            titleListNum = 4;
                        }
                        Regex regexPlan = new Regex(@"^Plan(?: [\w]+)? - (.+)$");
                        Match regexMath = regexPlan.Match(page);
                        if (regexPlan.IsMatch(page))
                        {
                            string currentLocation = regexMath.Groups[1].Value.Trim();
                            string firstWord = currentLocation.Split(' ').First();

                            if (firstWord != regexPlanBufferString)
                            {
                                CreateTitle("[" + firstWord + "]");
                                titleList.Add("["+firstWord+"]");
                                regexPlanBufferString = firstWord;

                                regexPlanIndexColor++;
                                if (regexPlanIndexColor >= gradient.Count)
                                    regexPlanIndexColor = 0;
                            }

                            listView1.SmallImageList.Images.Add(CreateColorIcon(gradient[regexPlanIndexColor]));
                        }
             

                        if (titleListNum != oldTitleNum)
                        {
                            CreateTitle(titleList[titleListNum], titleListNum);

                            listView1.SmallImageList.Images.Add(CreateColorIcon(System.Drawing.Color.FromArgb(0, 253, 161, 60)));
                        }

                        if (titleListNum == 2)
                            listView1.SmallImageList.Images.Add(CreateColorIcon(System.Drawing.Color.FromArgb(255, 5, 112, 176)));
                        
                        else if (titleListNum == 3)
                            listView1.SmallImageList.Images.Add(CreateColorIcon(System.Drawing.Color.FromArgb(255, 253, 161, 60)));
                        

                        item.ImageIndex = listView1.SmallImageList.Images.Count - 1;


                        item.SubItems.Add(page);
                        item.SubItems.Add((string)namesArrayU.GetValue(index-1));       

                        listView1.Items.Add(item);


                    }));

            if (buffer)
            {
                // Восстанавливаем прокрутку
                if (listView1.Items.Count > 0)
                {
                    if (topIndex < listView1.Items.Count)
                    {
                        listView1.Invoke(new Action(() =>
                        {
                            listView1.TopItem = listView1.Items[topIndex];
                        }));
                        
                    }
                    else if (listView1.Items.Count > 0)
                    {
                        listView1.Invoke(new Action(() =>
                        {
                            listView1.TopItem = listView1.Items[listView1.Items.Count - 1];
                        }));
                        
                    }
                }

            }

            // listView1.Columns[0].Text = newText;
        }
        
        private void CreateTitle(string text)
        {
            var headerItem = new ListViewItem();

            headerItem.SubItems.Add(__CreateTitleText(text));
            headerItem.ForeColor = System.Drawing.Color.White;
            headerItem.BackColor = System.Drawing.Color.Gray;
            headerItem.Font = new System.Drawing.Font(listView1.Font, FontStyle.Bold);
            headerItem.Tag = "header";
            listView1.Items.Add(headerItem);
        }

        private void CreateTitle(string text, int num)
        {
            var headerItem = new ListViewItem("[" + num + "]");

            headerItem.SubItems.Add(__CreateTitleText(text));
            headerItem.ForeColor = System.Drawing.Color.White;
            headerItem.BackColor = System.Drawing.Color.Gray;
            headerItem.Font = new System.Drawing.Font(listView1.Font, FontStyle.Bold);
            headerItem.Tag = "header";
            listView1.Items.Add(headerItem);
        }
        

        private string __CreateTitleText(string text, int totalWidth = 30, char paddingChar = '-') 
        {
            if (string.IsNullOrEmpty(text))
                return new string(paddingChar, totalWidth);

            text = text.Trim();

            int totalPadding = totalWidth - text.Length;
            if (totalPadding < 0)
                return text;

            int leftPadding = totalPadding / 2;
            int rightPadding = totalPadding - leftPadding;

            return new string(paddingChar, leftPadding) + " "+text+" " + new string(paddingChar, rightPadding);
        }

        private void listView1_Click(object sender, EventArgs e)
        {

        }

        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
          
        }

        private List<ListViewItem> SearchItem(string searchText, System.Windows.Forms.ListView obj)
        {
            List<ListViewItem> listViewItems = new List<ListViewItem>();

            foreach (ListViewItem item in obj.Items)
            {
                // Сбрасываем цвет для всех элементов
                if (titleList.Any(title => item.SubItems[1].Text.Contains(title)))
                    continue;

                if (string.IsNullOrEmpty(searchText))
                    continue;

                if(item.SubItems[0].Text == searchText)
                {
                    listViewItems.Add(item);
                }
                // Подсвечиваем совпадающие элементы
                if (item.SubItems[1].Text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Доабвляем совпадение
                    listViewItems.Add(item);
                }
                else
                {
                    try
                    {
                        // Если у нас тире, а не перечисление
                        Regex regex = new Regex(@"(G|L|g|l)(\d+)\s*-\s*(G|L|g|l)(\d+)");
                        Match match = regex.Match(item.SubItems[1].Text);

                        if (match.Success)
                        {
                            int firstNumber = int.Parse(match.Groups[2].Value); // 29
                            int secondNumber = int.Parse(match.Groups[4].Value); // 1

                            for (int i = firstNumber; i < secondNumber; i++)
                            {
                                // Нашли совпавдение между числами
                                if (searchText == ('G' + i.ToString()) || searchText == ('L' + i.ToString()))
                                {
                                    // Доабвляем совпадение
                                    listViewItems.Add(item);
                                }
                            }

                        }
                        // Если с точкой
                        else if (searchText.Contains('.'))
                        {

                            regex = new Regex(@"(G|L|g|l)(\d+)\.(\d+)\s*-\s*(G|L|g|l)?(\d+)\.(\d+)");
                            match = regex.Match(item.SubItems[1].Text);

                            if (match.Success)
                            {
                                int firstNumber1 = int.Parse(match.Groups[2].Value);
                                int firstNumber2 = int.Parse(match.Groups[3].Value);

                                int secondNumber1 = int.Parse(match.Groups[5].Value);
                                int secondNumber2 = int.Parse(match.Groups[6].Value);

                                for (int i = firstNumber2; i < secondNumber2; i++)
                                {
                                    // Нашли совпавдение между числами
                                    if (searchText == ('G' + firstNumber1.ToString() + '.' + i.ToString()) || searchText == ('L' + firstNumber1.ToString() + '.' + i.ToString()))
                                    {
                                        // Доабвляем совпадение
                                        listViewItems.Add(item);
                                    }
                                }

                            }

                        }
                    }
                    catch { Debug.WriteLine("Ошибка в " + item.SubItems[1].Text); return listViewItems; }
                }


            }
            return listViewItems;
        }

        private void HighlightItemsColorReset(System.Windows.Forms.ListView obj)
        {
            // Сбрасываем цвета
            foreach (ListViewItem item in obj.Items)
            {
                // Сброс цвета 
                if (titleList.Any(title => item.SubItems[1].Text.Contains(title)))
                    continue;

                item.BackColor = SystemColors.Window;
                item.ForeColor = SystemColors.WindowText;
            }
        }

        private void HighlightItemsInList(List<ListViewItem> listViewItems, System.Windows.Forms.ListView obj, Color back, Color fore)
        {
            listView1.BeginUpdate();



            foreach (ListViewItem item in listViewItems)
            {
                item.BackColor = back;
                item.ForeColor = fore;
            }
            listView1.EndUpdate();
        }

        private void SearchandUpdateCustomWindows()
        {
            List<ListViewItem> result = SearchItem(textBox1.Text, listView1);

            if (result.Count == 0 && string.IsNullOrEmpty(textBox1.Text))
            {
                customWindow.Caption = "Проводник по документам";
                isFound = false;
                index = 0;
                HighlightItemsColorReset(listView1);
                HighlightItemsInList(result, listView1, Color.Yellow, Color.Black);

            }
            else if (result.Count == 0 && !string.IsNullOrEmpty(textBox1.Text))
            {
                customWindow.Caption = "Найдено: " + 0 + "/" + 0;
                isFound = false;
                index = 0;
                HighlightItemsColorReset(listView1);
                HighlightItemsInList(result, listView1, Color.Yellow, Color.Black);

            }
            else
            {
                customWindow.Caption = "Найдено: " + (index + 1) + "/" + result.Count;
                result[0].EnsureVisible();
                isFound = true;
                index = 0;
                publicListViewItemsForEnterKey = result;
                listView1.TopItem = result[0];
                HighlightItemsColorReset(listView1);
                HighlightItemsInList(result, listView1, Color.Yellow, Color.Black);

                result[0].BackColor = System.Drawing.Color.GreenYellow;

            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Ищем и подсветчиваем элементы в списке
            SearchandUpdateCustomWindows();
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {

      


        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (isFound)
            {
                // верх
                if (e.KeyCode == Keys.Up)
                {
                    publicListViewItemsForEnterKey[index].BackColor = System.Drawing.Color.Yellow;

                    index = index + 1 <= 1 ? publicListViewItemsForEnterKey.Count - 1 : index - 1;
                    customWindow.Caption = "Найдено: " + (index + 1) + "/" + publicListViewItemsForEnterKey.Count;
                    e.Handled = true;
                }
                // вниз
                else if (e.KeyCode == Keys.Down)
                {
                    publicListViewItemsForEnterKey[index].BackColor = System.Drawing.Color.Yellow;

                    index = index + 1 >= publicListViewItemsForEnterKey.Count ? 0 : index + 1;
                    customWindow.Caption = "Найдено: " + (index + 1) + "/" + publicListViewItemsForEnterKey.Count;
                    e.Handled = true;
                }

                if (e.KeyCode == Keys.Enter)
                {
                    string itemText = publicListViewItemsForEnterKey[index].SubItems[2].Text;
                    GoToPageByNameU(itemText);
                    e.SuppressKeyPress = true;
                }

                listView1.BeginUpdate();

                publicListViewItemsForEnterKey[index].EnsureVisible();
                listView1.TopItem = publicListViewItemsForEnterKey[index];
                publicListViewItemsForEnterKey[index].BackColor = System.Drawing.Color.GreenYellow;
                listView1.EndUpdate();
            }
            else if (textBox1.Focused && string.IsNullOrEmpty(textBox1.Text) && e.KeyCode == Keys.Down)
            {
                listView1.Focus();
                listView1.SelectedItems.Clear();
                listView1.TopItem.Focused = true;
                listView1.TopItem.Selected = true;
            }
        }

        private void listView1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && listView1.SelectedItems.Count == 1)
            {
                string itemText = listView1.SelectedItems[0].SubItems[2].Text;
                GoToPageByNameU(itemText);
            }
            if (e.KeyCode == Keys.R || (char)e.KeyCode == 'к' || (char)e.KeyCode == 'К')
            {
                customWindow.Caption = "Проводник по документам";
                UpdateExplorer(visioApp.ActiveDocument);
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            // Обработаем ситуации что бы не кликнули на стр
            if (listView1.SelectedItems.Count == 1)
            {
                var selectedItem = listView1.SelectedItems[0];
                string itemText = selectedItem.SubItems[2].Text;
                GoToPageByNameU(itemText);
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.SelectAll();
        }

        private void переименоватьF2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 1)
            {
                var selectedItem = listView1.SelectedItems[0];
                string newName = ShowRenameDialog(selectedItem.SubItems[1].Text);
                if(newName == null)
                {
                    return;
                }
                try
                {
                    Visio.Page targetPage = visioApp.ActiveDocument.Pages.ItemU[selectedItem.SubItems[2].Text];
                    targetPage.Name = newName;
                    selectedItem.SubItems[1].Text = targetPage.Name;
                    selectedItem.SubItems[2].Text = targetPage.NameU;

                }
                catch 
                {

                }


            }
        }


        public string ShowRenameDialog(string currentName)
        {
            using (var form = new Form())
            {
                form.Text = "Переименовать";
                form.Size = new Size(300, 120);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var textBox = new System.Windows.Forms.TextBox() { Text = currentName, Left = 20, Top = 20, Width = 240 };
                var buttonOk = new System.Windows.Forms.Button() { Text = "OK", Left = 130, Top = 50, Width = 60 };
                var buttonCancel = new System.Windows.Forms.Button() { Text = "Отмена", Left = 200, Top = 50, Width = 60 };
                var label = new System.Windows.Forms.Label() { Left = 20, Top = 3, Width = 240 };
                label.ForeColor = Color.Orange;

                buttonOk.DialogResult = DialogResult.OK;
                buttonCancel.DialogResult = DialogResult.Cancel;

                form.Controls.AddRange(new Control[] {textBox, buttonOk, buttonCancel, label });
                form.AcceptButton = buttonOk;
                form.CancelButton = buttonCancel;

                textBox.TextChanged += (object sender, EventArgs e) => 
                {
                    RenameDialogTextboxEvent(textBox, label, buttonOk);
                };
                form.Load += (object sender, EventArgs e) =>
                {
                    RenameDialogTextboxEvent(textBox, label, buttonOk);
                };

                var result = form.ShowDialog();

        


                if (result == DialogResult.OK)
                {
                    /*
                    foreach(ListViewItem item in listView1.Items)
                    {
                        if(textBox.Text == item.SubItems[2].Text)
                        {
                            MessageBox.Show("Такое имя уже существует",
                                "Ой",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return null;
                        }
                    }
                    */
                    return textBox.Text;
                }
            }

            return currentName;
        }

        private void RenameDialogTextboxEvent(System.Windows.Forms.TextBox textBox, System.Windows.Forms.Label label, System.Windows.Forms.Button button)
        {
            // Тут логика проверки 
            if (textBox == null || string.IsNullOrEmpty(textBox.Text))
            {
                label.Text = "";
                button.Enabled = false;
                return;
            }
            label.ForeColor = Color.Orange;
            button.Enabled = true;

            Regex deviceReg = new Regex(@"^G\d");
            Regex lightReg = new Regex(@"^L\d");

            if (textBox.Text[0] == ' ' || textBox.Text[textBox.Text.Length - 1] == ' ')
            {
                label.ForeColor = Color.OrangeRed;
                label.Text = "Имя содержит в конце или в начале пробел";
                button.Enabled = false;
            }
            else if (deviceReg.IsMatch(textBox.Text))
            {
                label.Text = "Это устройство";
                int count = DevicesCountRegex(textBox.Text);
                if(count != 0)
                {
                    label.Text += " " + count + " шт.";
                }
                

            }
            else if (lightReg.IsMatch(textBox.Text))
            {
                label.Text = "Это cвет";
                int count = DevicesCountRegex(textBox.Text);
                if (count != 0)
                {
                    label.Text += " " + count + " шт.";
                }
            }
            else
            {
                label.ForeColor = Color.Gray;
                label.Text = "Другое";
            }

        }

        private int DevicesCountRegex(string text)
        {
            Regex regex = new Regex(@"(G|L|g|l)(\d+)-(G|L|g|l)(\d+)");
            Match match = regex.Match(text);

            if (match.Success)
            {
                try
                {
                    int firstNumber = int.Parse(match.Groups[2].Value); 
                    int secondNumber = int.Parse(match.Groups[4].Value); 
                    return Math.Abs(firstNumber - secondNumber);
                }
                catch
                {
                    return 0;
                }
            }

            regex = new Regex(@"(G|L|g|l)(\d+)\.(\d+)-(G|L|g|l)?(\d+)\.(\d+)");
            match = regex.Match(text);

            if (match.Success)
            {
                try
                {
                    int firstNumber1 = int.Parse(match.Groups[2].Value);
                    int firstNumber2 = int.Parse(match.Groups[3].Value);
                    int secondNumber1 = int.Parse(match.Groups[5].Value);
                    int secondNumber2 = int.Parse(match.Groups[6].Value);

                    if (firstNumber1 - secondNumber1 != 0)
                        return 0;

                    return Math.Abs(firstNumber2 - secondNumber2);
                }
                catch
                {
                    return 0;
                }
            }
            return 0;
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    try
                    {
                        Visio.Page targetPage = visioApp.ActiveDocument.Pages.ItemU[item.SubItems[2].Text];
                        targetPage.Delete(1);
                       
                    }
                    catch
                    {

                    }
                }
                UpdateExplorer(visioApp.ActiveDocument);

            }
        }
    }


}
