using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Windows.Media;

namespace QueenMasterVisio.DeviceControl.Models
{
    internal class FormItem
    {
        public int Code { get; set; }
        public string Description { get; set; }
        public Visio.Shape Shape { get; set; }
        public Page Page { get; set; }
        public string Icon
        {
            get
            {
                switch (Code)
                {
                    case 1:
                        return "ℹ️";
                    case 2:
                        return "⚠️";
                    case 3:
                        return "❌";
                    default:
                        return " ";
                }
            }
        }

        public SolidColorBrush IconColor
        {
            get
            {
                switch (Code)
                {
                    case 1: return new SolidColorBrush(Colors.Blue);      
                    case 2: return new SolidColorBrush(Colors.Orange);   
                    case 3: return new SolidColorBrush(Colors.Red);     
                    default: return new SolidColorBrush(Colors.Black);
                }
            }
        }
        public int classError => int.Parse(Code.ToString()[0] + "");
        public string ShapeName => Shape?.NameU ?? "";
        public string PageName => Page?.Index.ToString() ?? "";
    }
}
