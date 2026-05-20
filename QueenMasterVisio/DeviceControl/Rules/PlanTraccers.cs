using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.DeviceControl.Results;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ValidationResult = QueenMasterVisio.DeviceControl.Results.ValidationResult;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class PlanTraccers : IRule
    {
        public string Name => "Проверка трассировок";

        public IEnumerable<ValidationResult> Validate(Visio.Page page, ValidationEngine engine)
        {
            if (page.IsPlanPage())
            {
                // Проверяем на уникальность
                List<string> duplicates = new List<string>();
                string resultStr = "";
                int shapeCount = 0;
                foreach (Visio.Shape shape in engine.Shapes)
                {
                    if (shape.Name.Contains("Device"))
                    {
                        shapeCount++;

                        string gName = Tools.CellFormulaGet(shape, "Prop.Number.Value");
                        string gIntName = gName.Replace("\"", "").Replace(" ", "");
                        bool check = false;
                        foreach (string item in duplicates)
                        {
                            if (item == gName)
                            {
                                // Уже есть такой, не уникален
                                check = true;
                                resultStr = "G" + gIntName + " обнаружен дубликат";
                                yield return new ValidationResult
                                {
                                    RuleName = Name,
                                    TargetPage = page,
                                    TargetShape = shape,
                                    Severity = ResultSeverity.Error,
                                    Message = resultStr,
                                };
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
                            resultStr = "G" + gIntName + " не хватает трассеров - ";
                            List<string> array = new List<string>();
                            foreach (string wireName in result)
                            {
                                array.Add(wireName);
                            }
                            resultStr += string.Join(", ", result);

                            yield return new ValidationResult
                            {
                                RuleName = Name,
                                TargetPage = page,
                                TargetShape = shape,
                                Severity = ResultSeverity.Error,
                                Message = resultStr,
                            };
                        }
                        result = PlanDeviceCheckConnectsExcess(shape);
                        if (result.Count > 0)
                        {

                            resultStr = "G" + gIntName + " перетрассирован - ";
                            List<string> array = new List<string>();

                            foreach (string wireName in result)
                            {
                                array.Add(wireName);
                            }
                            resultStr += string.Join(", ", result);

                            yield return new ValidationResult
                            {
                                RuleName = Name,
                                TargetPage = page,
                                TargetShape = shape,
                                Severity = ResultSeverity.Error,
                                Message = resultStr,
                            };

                        }
                    }
                }


            }
        }


        // Получаем количества ожидаемых подключений по каждому девайсу
        // Так например, можно определить количество rpi или бордов
        public static Dictionary<string, int> GetCountAllExpectedConnects(Visio.Page page)
        {
            Dictionary<string, int> result = new Dictionary<string, int>();

            // Заполняем
            foreach (Wire wire in WireService.wires.Where(w => w.isWire))
            {
                if (!wire.name.Contains("Other"))
                    result.Add(wire.name, 0);
            }
            // Считаем
            foreach (Visio.Shape shape in page.Shapes)
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
                    if (connectedShape.IsLine())
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
                if (connectedShape.IsLine())
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

    }
}
