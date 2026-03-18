using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio
{
    internal class PageElement
    {
        public string name;
        public Page page;
        public string comment;
        public string type;
        
        public PageElement(string _name, Page _page, string _comment, string _type)
        {
            name = _name;
            page = _page;
            comment = _comment;
            type = _type;
        }
    }
}
