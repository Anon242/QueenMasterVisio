using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio
{
    internal class Wire
    {
        public int id;
        public string name;
        public bool isWire;
        public string comment;
        public string color;
        public bool arrow;
        public string defaultCable;
        public string voltage;

        public Wire(int _id, string _name, bool _isWire, string _comment, string _color, bool _arrow, string _defaultCable, string _voltage)
        {
            id = _id;
            name = _name;
            isWire = _isWire;
            comment = _comment;
            color = _color;
            arrow = _arrow;
            defaultCable = _defaultCable;
            voltage = _voltage;
        }
    }
}
