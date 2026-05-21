using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.DeviceControl.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ValidationResult = QueenMasterVisio.DeviceControl.Results.ValidationResult;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class PageName : IRule
    {
        public string Name => "Проверка имени";

        public IEnumerable<ValidationResult> Validate(Visio.Page page, ValidationEngine engine)
        {
            if (!page.IsPlanPage() && !page.IsAutoTracePage())
            {
                /*
                if (!(new Regex(@"^G\d").IsMatch(page.Name) || new Regex(@"^L\d").IsMatch(page.Name)))
                {
                    yield return new ValidationResult
                    {
                        RuleName = Name,
                        TargetPage = page,
                        TargetShape = null,
                        Severity = ResultSeverity.Warning,
                        Message = "Неопределенное название страницы",
                    };
                }
                */

                if(page.Name.Length >= 65)
                {
                    yield return new ValidationResult
                    {
                        RuleName = Name,
                        TargetPage = page,
                        TargetShape = null,
                        Severity = ResultSeverity.Warning,
                        Message = "Слишком длинное название",
                    };
                }

                // ПРОВЕРКА НА РУССКИЕ НАЗВАНИЯ
            }
        }
    }
}
