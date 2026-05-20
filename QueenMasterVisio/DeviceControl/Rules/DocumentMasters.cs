using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.DeviceControl.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Shapes;
using ValidationResult = QueenMasterVisio.DeviceControl.Results.ValidationResult;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.DeviceControl.Rules
{
    internal class DocumentMasters : IRule
    {
        public string Name => "Проверка клемм";

        public IEnumerable<ValidationResult> Validate(Visio.Page page, ValidationEngine engine)
        {
           var masters = engine.Shapes[1].Document.Masters;
            List <string> mastersStr = new List<string>();
            foreach (Master item in masters)
            {
                if(!item.NameU.Contains("Master"))
                    mastersStr.Add(item.NameU);
            }
            List<string> dup = GetAllDuplicateStrings(mastersStr);
            if (dup.Count > 0) 
            {
            foreach (string item in dup)
             {
                    if(!item.Contains("."))
                    yield return new ValidationResult
                    {
                        RuleName = Name,
                        TargetPage = null,
                        TargetShape = null,
                        Severity = ResultSeverity.Warning,
                        Message = "Обнаружен дубликат мастера в структуре документа: " + item,
                    };
                }
            }
        }

        public static List<string> GetAllDuplicateStrings(List<string> source)
        {
            string Normalize(string s) => Regex.Replace(s, @"\.\d+$", "");

            return source
                .GroupBy(Normalize)
                .Where(g => g.Count() > 1)
                .SelectMany(g => g)      // разворачиваем все строки из групп
                .ToList();
        }
    }
}
