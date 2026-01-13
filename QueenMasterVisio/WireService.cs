using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace QueenMasterVisio
{
    internal class WireService
    {
        public static readonly List<Wire> wires = new List<Wire>()
        {
            new Wire(0,"Plan",    false,"Plan","THEMEGUARD(RGB(0,0,0))",                    false, "",          ""          ),
            new Wire(1,"Px",      true, "230v wires","THEMEGUARD(RGB(255,0,0))",            false, "3 x 2.5",   "230v"      ),
            new Wire(2,"Ex",      true, "Ethernet wires","THEMEGUARD(RGB(0,176,240))",      false, "UTP Cat 5E","5v"        ),
            new Wire(3,"Rx",      true, "RS485 wires","THEMEGUARD(RGB(112,48,160))",        false, "UTP Cat 5E","12v"       ),
            new Wire(4,"Dx",      true, "12v wires","THEMEGUARD(RGB(255,235,0))",           false, "3 x 1.5",   "12v"       ),
            new Wire(5,"Cx",      true, "12v signal wires","THEMEGUARD(RGB(255,192,0))",    true,  "3 x 1.5",    "12v"      ),
            new Wire(6,"Sx",      true, "Sensor wires","THEMEGUARD(RGB(0,176,80))",         true,  "UTP Cat 5E", "12v"      ),
            new Wire(7,"Yx",      true, "Emergency buttons","THEMEGUARD(RGB(185,185,185))", false, "3 x 0.75",   "12v"      ),
            new Wire(8,"Vx",      true, "CCTV wires","THEMEGUARD(RGB(204,194,217))",        false, "UTP Cat 5E","36v-57v"   ),
            new Wire(9,"Lx",      true, "Light wires","THEMEGUARD(RGB(0,32,96))",           false, "3 x 1.5",   "12v"       ),
            new Wire(10,"Ax",     true, "Sound wires","THEMEGUARD(RGB(234,112,13))",        false, "2 x 1.5",   "24v"       ),
            new Wire(11,"Other1", true, "Other1 wires","THEMEGUARD(RGB(0,0,0))",            false, "",          ""          ),
            new Wire(12,"Other2", true, "Other2 wires","THEMEGUARD(RGB(0,0,0))",            false, "",          ""          ),
            new Wire(13,"Develop",false,"","THEMEGUARD(RGB(0,0,0))",                        false, "",          ""          ),
        };


        // Проверяет, является ли wire.name трасером
        public static bool ThatIsTraccer(string name)
        {
            return wires.Any(w => w.name == name && w.isWire);
        }
        // Получить объект трасера зная лишь имя
        public static Wire GetWireByName(string name)
        {
            return wires.FirstOrDefault(w => w.name == name);
        }
        // Получить объект трасера зная лишь цвет
        public static Wire GetWireByColorName(string colorName)
        {
            return wires.FirstOrDefault(w => w.color == colorName);
        }

    }
}
