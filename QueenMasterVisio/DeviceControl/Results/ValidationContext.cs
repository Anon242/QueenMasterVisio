using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.DeviceControl.Results
{
    public class ValidationContext
    {
        public double ToleranceMm { get; set; } = 0.5;
        public bool IncludeHiddenShapes { get; set; } = false;
        // можно добавить другие настройки
    }
}
