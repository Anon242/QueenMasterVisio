using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Visio.Page;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Reflection.Emit;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.ObjectModel;
using System.Collections;


namespace QueenMasterVisio
{
    // Это класс к которому обращается BackroundWorker и главный класс
    // Он содержит методы проверок страниц, соеденений и тд
    // Каждый тест должен возращать List своих значений
    static internal class Checker
    {
        public static void CheckDevicesInPlan(Page page)
        {
            // Проверяем на уникальность
            List<string> duplicates = new List<string>();
            string a = "";
            int shapeCount = 0;
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.Name.Contains("Device"))
                {
                    shapeCount++;

                    string gName = Tools.CellFormulaGet(shape, "Prop.Number.Value");

                    bool check = false;
                    foreach (string item in duplicates)
                    {
                        if (item == gName)
                        {
                            // Уже есть такой, не уникален
                            check = true;
                            a += gName + " ЕСТЬ СОВПАДЕНИЕ" + '\n';
                            break;
                        }

                    }
                    if (!check)
                    {
                        // Уникален
                        duplicates.Add(gName);
                    }
                    List<string> result = PlanDeviceCheckConnects(shape);
                    if (result.Count > 0)
                    {
                        a += gName + " не хватает:  " ;
                        foreach (string wireName in result)
                        {
                            a += wireName + ", ";
                        }
                        a += "\n";
                    }
                    result = PlanDeviceCheckConnectsExcess(shape);
                    if (result.Count > 0)
                    {
                        a += gName + " перетрасирован с:  " ;
                        foreach (string wireName in result)
                        {
                            a += wireName + ", ";
                        }
                        a += "\n";

                    }
                }
            }
            a+= "Всего ожидаемых подключений: "+"\n";
            foreach(KeyValuePair<string, int> kp in GetCountAllExpectedConnects(page))
            {
                a += kp.Key + ": " + kp.Value + '\n';
            }
            if (a != "")
                MessageBox.Show(a, "Всего устройств: "+ shapeCount, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    
        // Получаем количества ожидаемых подключений по каждому девайсу
        // Так например, можно определить количество rpi или бордов
        public static Dictionary<string, int> GetCountAllExpectedConnects(Visio.Page page)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();

            // Заполняем
            foreach (Wire wire in WireService.wires.Where(w => w.isWire))
            {
                if(!wire.name.Contains("Other"))
                    result.Add(wire.name, 0);
            }
            // Считаем
            foreach(Visio.Shape shape in page.Shapes)
            {
                if (shape.Name.Contains("Device"))
                {
                    foreach (Wire wire in GetExpectedConnectionsInDevice(shape))
                    {
                        result[wire.name]++;
                    }
                }
            }
            return result;
        }
        // Получаем ожидаемые трасеры в устройство
        private static List<Wire> GetExpectedConnectionsInDevice(Visio.Shape shape)
        {
            List<Wire> expectedConnectionList = new List<Wire>();

            // Записываем ожидаемые подключения и игнорим Yx
            foreach (Wire wire in WireService.wires.Where(w => w.isWire))
            {
                // Проверка на существование Cell
                if (Tools.CellExistsCheck(shape, "Prop." + wire.name))
                {
                    string value = Tools.CellFormulaGet(shape, "Prop." + wire.name + ".Value");
                    if (value.Split('(')[1].Split(',')[0] == "0")
                        expectedConnectionList.Add(wire);
                }

            }
            return expectedConnectionList;
        }
    	
		// Чекаем коннекты у Device на плане
		// Возращает несовпавшие
		private static List<string> PlanDeviceCheckConnects(Visio.Shape shape)
        {
            // Если не девайс выходим
            if (!shape.Name.Contains("Device"))
                return new List<string>();

            // Имя устройства
            string gName = Tools.CellFormulaGet(shape, "Prop.Number.Value");

            List<Wire> expectedConnectionList = GetExpectedConnectionsInDevice(shape).Where(w => w.name != "Yx").ToList();

            // Мы получили Wires которые мы ожидаем быть подключенными
            // Теперь сравниваем с действительностью

            List<string> result = new List<string>();

            foreach (Wire wire in expectedConnectionList)
            {
                bool isFindWire = false;
                foreach (Connect connect in shape.FromConnects)
                {
                    Visio.Shape connectedShape = connect.FromSheet;
                    if (isLine(connectedShape))
                    {
                        // Если найдено
                        if (Tools.CellFormulaGet(connectedShape, "LineColor") == wire.color)
                        {
                            isFindWire = true;
                            break;
                        }
                    }

                }
                if (!isFindWire)
                {
                    result.Add(wire.name);
                }
            }

            return result.Distinct().ToList();
        }

        private static List<string> PlanDeviceCheckConnectsExcess(Visio.Shape shape)
        {
            // Если не девайс выходим
            if (!shape.Name.Contains("Device"))
                return new List<string>();

            // Имя устройства
            string gName = Tools.CellFormulaGet(shape, "Prop.Number.Value");

            // Получаем ожидаемые подключения, исключаем Yx
            List<Wire> expectedConnectionList = GetExpectedConnectionsInDevice(shape).Where(w => w.name != "Yx").ToList();
            List<Wire> excessConnectionList = WireService.wires.Where(w => w.isWire).Except(expectedConnectionList).ToList().Where(w => w.name != "Yx").ToList();
            // Мы получили Wires которые мы ожидаем быть подключенными
            // Теперь сравниваем с действительностью

            List<string> result = new List<string>();

            foreach (Connect connect in shape.FromConnects)
            {
                Visio.Shape connectedShape = connect.FromSheet;
                if (isLine(connectedShape))
                {
                    foreach (Wire wire in excessConnectionList)
                    {
                        // Если найдено
                        if (Tools.CellFormulaGet(connectedShape, "LineColor") == wire.color)
                        {
                            result.Add(wire.name);
                            break;
                        }

                    }
                }
            }

            return result.Distinct().ToList();
        }

        // Проверяем что это линия
        public static bool isLine(Visio.Shape shape)
        {
            if (shape?.Name != null)
                return shape.Name.Contains("Динамическая соединительная линия");
            else return false;
        }
        /*
         * Проверяет существования Cell
         */



    }
}
