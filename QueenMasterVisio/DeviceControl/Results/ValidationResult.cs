using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.DeviceControl.Results
{
    public enum ResultSeverity { Info = 1, Warning = 2, Error = 3 }

    public class ValidationResult
    {
        public string RuleName { get; set; }
        public string Message { get; set; }
        public ResultSeverity Severity { get; set; }
        public Microsoft.Office.Interop.Visio.Shape TargetShape { get; set; }
        public Microsoft.Office.Interop.Visio.Page TargetPage { get; set; }

    }
}
