using QueenMasterVisio.Core.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Shapes;

namespace QueenMasterVisio.Core.Models
{
    internal class Cable
    {
        Microsoft.Office.Interop.Visio.Shape shape;
        public string type;
        public string code;
        public string deviceName;
        public Cable(Microsoft.Office.Interop.Visio.Shape shape)
        {
            this.shape = shape;
            try
            {
                type = shape.GetCellResultString("Prop.type");
                code = shape.GetCellResultString("Prop.Name"); 
                deviceName = shape.GetCellResultString("Prop.Device");
            }
            catch (Exception)
            {

                
            }

        }
    }
}
