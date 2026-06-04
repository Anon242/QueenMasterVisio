using QueenMasterVisio.DeviceControl.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;
using Shape = Microsoft.Office.Interop.Visio.Shape;
using Page = Microsoft.Office.Interop.Visio.Page;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Models;
using System.Diagnostics;
using QueenMasterVisio.DeviceControl.Service;
using System.Runtime.ConstrainedExecution;
using System.Windows.Documents;

namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class Cables : IRule
    {
        public string Name => "Проверка кабелей";

        public IEnumerable<ValidationResult> Validate(Page page, ValidationEngine engine)
        {

            // Требуется список девайсов, если нет отклоняем
            string userPageCode = page.GetUserPageCode();
            if(userPageCode != "Device" && userPageCode != "Light")
            {
                yield return new ValidationResult
                {
                    RuleName = Name,
                    TargetPage = page,
                    TargetShape = null,
                    Severity = ResultSeverity.Error,
                    Message = Name +": Страница не является девайсом, возможно она не привязана к объекту плана",
                };
                yield break;
            }
            string [] deviceList = page.GetCellFormulaU("User.deviceList").Split(';');
            if (deviceList.Length <= 0)
            {
                yield return new ValidationResult
                {
                    RuleName = Name,
                    TargetPage = page,
                    TargetShape = null,
                    Severity = ResultSeverity.Error,
                    Message = Name + ": Отсутсвует привязка к объекту, привяжите страницу к объекту плана",
                };
                yield break;
            }


            List<Shape> cables = _GetCables(engine);
            // Теперь получим откуда это устройство

            foreach (var _cable in cables)
            {
                Cable cable = new Cable(_cable);

                /*
                // Чекаем версию если версии загружены
                VersionsList versions = new VersionsList(ThisAddIn.links.VersionsFile);

                if (versions.Versions.Any(a => a.Item1 == "Cable"))
                {
                    var ver = versions.Versions.FirstOrDefault(a => a.Item1 == "Cable");
                    string cableVersion = _cable.GetCellResultString("User.Version");
                    if (!string.IsNullOrEmpty(cableVersion))
                    {
                        if (float.Parse(cableVersion) < ver.Item2)
                        {
                            yield return new ValidationResult
                            {
                                RuleName = Name,
                                TargetPage = page,
                                TargetShape = _cable,
                                Severity = ResultSeverity.Warning,
                                Message = "Кабель на странице: " + cableVersion + " последняя версия кабеля: " + ver.Item2,
                            };
                            yield break;
                        }
                    }
                    else if(!_cable.Name.Contains("Cable.440"))
                    {
                        yield return new ValidationResult
                        {
                            RuleName = Name,
                            TargetPage = page,
                            TargetShape = _cable,
                            Severity = ResultSeverity.Warning,
                            Message = "Образец кабеля старой версии, необходимо его заменить",
                        };
                        yield break;
                    }

                }
                */
                if (!string.IsNullOrEmpty(cable.deviceName) && !string.IsNullOrEmpty(cable.type))
                {

                    Array gluedShapes = _cable.GluedShapes((short)Visio.VisGluedShapesFlags.visGluedShapesAll1D, "", null);
                    if (gluedShapes.Length == 0 && !cable.code.Contains("R"))
                    {
                        yield return new ValidationResult
                        {
                            RuleName = Name,
                            TargetPage = page,
                            TargetShape = _cable,
                            Severity = ResultSeverity.Warning,
                            Message = "Кабель не подключен",
                        };
                    }


                    // Еще проверка наследуется ли он от правильного мастера

                    if (cable.deviceName.Length <= 1 || cable.code.Length <= 1)
                    {
                        yield return new ValidationResult
                        {
                            RuleName = Name,
                            TargetPage = page,
                            TargetShape = _cable,
                            Severity = ResultSeverity.Error,
                            Message = "Кабель имеет невалидные значения: " + cable.deviceName + " || " + cable.code,
                        };
                    }
                    else
                    {
                        bool isFinded = false;
                        // Пробуем найти
                        if (cable.deviceName.Contains("J"))
                        {
                            yield return new ValidationResult
                            {
                                RuleName = Name,
                                TargetPage = page,
                                TargetShape = _cable,
                                Severity = ResultSeverity.Info,
                                Message = "Кабель " + cable.deviceName + " " + cable.type + " " + cable.code + " не может быть опознан, но учтен",
                            };
                            isFinded = true;
                            goto Finish;
                        }
                        if(cable.deviceName.Contains("G") || cable.deviceName.Contains("L"))
                        foreach (Page _page in page.Document.Pages)
                        {
                            if (_page.Name.Contains(cable.deviceName + " "))
                            {

                                List<Shape> _cables = _GetCables(_page);
                                foreach (Shape _shape in _cables)
                                {
                                    Cable newcable = new Cable(_shape);
                                    if (cable.code == newcable.code && cable.type == newcable.type)
                                    {
                                        yield return new ValidationResult
                                        {
                                            RuleName = Name,
                                            TargetPage = page,
                                            TargetShape = _cable,
                                            Severity = ResultSeverity.Info,
                                            Message = "Кабель " + cable.deviceName + " " + cable.type + " " + cable.code + " найден на странице: " + _page.Name.Replace("\"", ""),
                                        };
                                        isFinded = true;
                                        goto Finish;
                                    }
                                }
                            }
                        }

                    Finish:

                        if (!isFinded)
                        {
                            yield return new ValidationResult
                            {
                                RuleName = Name,
                                TargetPage = page,
                                TargetShape = _cable,
                                Severity = ResultSeverity.Error,
                                Message = "Ответный кабель не найден или не найдена страница: " + cable.deviceName,
                            };
                        }

                    }
                }
            }
    
        }

        private List<Shape> _GetCables(ValidationEngine engine)
        {
            List<Shape> shapes = new List<Shape>();
            foreach (Shape shape in engine.Shapes)
            {
                // Обычная клемма
                if (shape.Name.Contains("Cable"))
                {
                    shapes.Add(shape);
                }
            }
            return shapes;
        }

        private List<Shape> _GetCables(Page engine)
        {
            List<Shape> shapes = new List<Shape>();
            foreach (Shape shape in engine.Shapes)
            {
                // Обычная клемма
                if (shape.Name.Contains("Cable"))
                {
                    shapes.Add(shape);
                }
            }
            return shapes;
        }
    }

}


/*
 yield return new ValidationResult
                                {
                                    RuleName = Name,
                                    TargetPage = page,
                                    TargetShape = shape,
                                    Severity = ResultSeverity.Error,
                                    Message = resultStr,
                                };
*/