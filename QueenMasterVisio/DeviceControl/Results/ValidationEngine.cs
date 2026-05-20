using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.DeviceControl.Results
{
    public class ValidationEngine
    {
        public List<Visio.Shape> Shapes { get; set; }
        
        public ValidationEngine(Visio.Page page) {
            if (page == null)
                return;

            Shapes = _GetAllShapes(page);
        }
        private List<Visio.Shape> _GetAllShapes(Visio.Page page)
        {
            var allShapes = new List<Visio.Shape>();
            _AddShapesRecursive(page.Shapes, allShapes);
            return allShapes;
        }

        private void _AddShapesRecursive(Visio.Shapes shapes, List<Visio.Shape> list)
        {
            foreach (Visio.Shape shape in shapes)
            {
                list.Add(shape);
                if (shape.Type == (short)Visio.VisShapeTypes.visTypeGroup)
                {
                    _AddShapesRecursive(shape.Shapes, list);
                }
            }
        }
    }
}
