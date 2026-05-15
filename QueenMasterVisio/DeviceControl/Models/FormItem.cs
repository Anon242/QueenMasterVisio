using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.DeviceControl.Models
{
    internal class FormItem
    {
        public int Code { get; set; }
        public string Description { get; set; }
        public string Shape { get; set; }
        public int Page { get; set; }
        public System.Windows.Media.ImageSource Icon { get; set; }
        public int classError => int.Parse(Code.ToString()[0] + "");
    }
}
