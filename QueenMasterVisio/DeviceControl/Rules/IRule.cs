using QueenMasterVisio.DeviceControl.Results;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using ValidationResult = QueenMasterVisio.DeviceControl.Results.ValidationResult;

namespace QueenMasterVisio.DeviceControl.Rules
{
    public interface IRule
    {
        string Name { get; }
        IEnumerable<ValidationResult> Validate(Microsoft.Office.Interop.Visio.Page page, ValidationEngine engine);

    }
}
