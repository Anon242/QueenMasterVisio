using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.DeviceControl.Results;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio = Microsoft.Office.Interop.Visio;
namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class Terminals : IRule
    {
        public string Name => "Проверка клемм";

        public IEnumerable<ValidationResult> Validate(Microsoft.Office.Interop.Visio.Page page, ValidationEngine engine)
        {
            List<string> names = new List<string>() { "TerminalScrew", "TerminalShield"};
            foreach (Shape shape in engine.Shapes)
            {
                // Обычная клемма
                if (shape.NameU.Contains(names[0])) 
                {
                    Array gluedShapes = shape.GluedShapes((short)Visio.VisGluedShapesFlags.visGluedShapesAll1D,"",null);
                    if (gluedShapes.Length == 0)
                    {
                        yield return new ValidationResult
                        {
                            RuleName = Name,
                            TargetPage = page,
                            TargetShape = shape,
                            Severity = ResultSeverity.Warning,
                            Message = "Клемма не подключена",
                        };
                    }
                }
                // Клемма щит 1,3
                else if (shape.NameU.Contains(names[1]))
                {
                    Array gluedShapes = shape.GluedShapes((short)Visio.VisGluedShapesFlags.visGluedShapesAll1D, "", null);

                    if (gluedShapes.Length == 0)
                    {
                        yield return new ValidationResult
                        {
                            RuleName = Name,
                            TargetPage = page,
                            TargetShape = shape,
                            Severity = ResultSeverity.Warning,
                            Message = "Клемма не подключена",
                        };
                    }
                }
            }
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