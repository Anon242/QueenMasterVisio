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


namespace QueenMasterVisio.DeviceCheck
{
    public partial class DeviceCheck : Form
    {
        Visio.Application app;
        public DeviceCheck(Visio.Application _application)
        {
            app = _application;
            InitializeComponent();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Thread th = new Thread(() => { 
            foreach (Visio.Page page in app.ActiveDocument.Pages)
            {
                string []arr = extractGValues(page.NameU);
                if(arr.Length != 0)
                {
                    bool res = LookDevicesOnPlan(app.ActivePage);
                        listBox1.Invoke(new Action(() =>
                        {
                            listBox1.Items.Add((res ? "" : "---") + page.NameU);
                        }));
                        
                }
            }
            });
            th.IsBackground = true;
            th.Start();
        }

        public bool LookDevicesOnPlan(Visio.Page page)
        {
            KeyValuePair<Visio.Page, Visio.Shape> dev = SearchDeviceOnDocument(page);
            if (dev.Key == null)
                return false;

            string[] pageNameCodeArray = extractGValues(page.Name);
            string pageNameCode = pageNameCodeArray[0].Replace("G", "");
            string result = "";

            Visio.Page targetPage = dev.Key;

            // Получили страницу плана, теперь получим все соединения 
            if (targetPage != null)
            {
                string text = CableSchedule.Generate(targetPage);

                // Вытщаим из таблицы только строки с нашим девайсом
                foreach (string col in text.Split('\n'))
                {
                    if ((col.Contains("G" + pageNameCode + ";") || col.Contains("G" + pageNameCode + " ")))
                    {
                        result += col + '\n';
                    }
                }
                targetPage = null;
            }
            // Ищем кол во кабелей на странице
            // Еще надо проверить что он на странице, а не за границей
            int countShapes = 0;
            foreach (Visio.Shape shape in page.Shapes)
            {

                if (shape.NameU.Contains("Cable"))
                {
                    countShapes++;
                }

                else if (shape.Type == (short)Visio.VisShapeTypes.visTypeGroup)
                {
                    foreach (Visio.Shape shapein in shape.Shapes)
                    {
                        //Debug.WriteLine(shapein.NameU);

                        if (shapein.NameU.Contains("Cable"))
                        {
                            countShapes++;
                        }

                    }
                }
            }

            //MessageBox.Show("Ожидаемое количество: " + (result.Split('\n').Length - 1) + '\n' + "На странице: " + countShapes, "Проверка кабелей", MessageBoxButtons.OK);

            //Debug.WriteLine(result);
           // Debug.WriteLine(countShapes);
            return (result.Split('\n').Length - 1) == countShapes;
        }

        public KeyValuePair<Visio.Page, Visio.Shape> SearchDeviceOnDocument(Visio.Page page)
        {
            KeyValuePair<Visio.Page, Visio.Shape> result = new KeyValuePair<Visio.Page, Visio.Shape>(null, null);
            string[] pageNameCodeArray = extractGValues(page.Name);
            // Проверяем что это вообще девайс и получаем его G код
            if (pageNameCodeArray.Length == 0)
            {
                //MessageBox.Show("Алгоритм не определил сигнатуру устройства", "Страница не распознана", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return result;
            }

            string pageNameCode = pageNameCodeArray[0].Replace("G", "");

            // Сначала ищем и получаем план в котором будет находится наш номер
            foreach (Visio.Page searchPage in page.Document.Pages)
            {
                if (Tools.CellExistsCheck(searchPage, "User.pageCode"))
                {
                    // Нашли план
                    if (Tools.CellFormulaGet(searchPage, "User.pageCode") == "Plan")
                    {
                        // Ищем объект девайса
                        foreach (Visio.Shape shape in searchPage.Shapes)
                        {
                            if (!shape.Name.Contains("Device") && !shape.Name.Contains("Shield")) continue;

                            // Если существует 
                            if (Tools.CellExistsCheck(shape, "Prop.Number"))
                            {
                                string nameValue = shape.CellsU["Prop.Number"].FormulaU.Replace("\"", "");
                                Debug.WriteLine(nameValue);
                                // Нашли совпадение
                                if (nameValue == pageNameCode)
                                {
                                    result = new KeyValuePair<Visio.Page, Visio.Shape>(searchPage, shape);
                                    Debug.WriteLine("НАШЛИ nameValue == pageNameCode: " + nameValue == pageNameCode);
                                    return result;

                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
        private static string[] extractGValues(string input)
        {
            Regex regex = new Regex(@"G\d+(?:\.\d+)?\b");

            MatchCollection matches = regex.Matches(input);

            List<string> result = new List<string>();
            foreach (Match match in matches)
            {
                result.Add(match.Value);
            }

            return result.ToArray();
        }

    }
}
