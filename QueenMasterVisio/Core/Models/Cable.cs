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
        public string way;
        public string deviceName;
        public bool isAutoCable = false;
        public Cable(Microsoft.Office.Interop.Visio.Shape shape)
        {
            this.shape = shape;
            try
            {
                type = shape.GetCellResultString("Prop.type");
                way = shape.GetCellResultString("Prop.Name"); 
                deviceName = shape.GetCellResultString("Prop.Device");
                if (shape.GetCellFormulaU("Prop.To").Contains("IFERROR(Pages["))
                    isAutoCable = true;
            }
            catch (Exception)
            {

                
            }

        }
    }
}
