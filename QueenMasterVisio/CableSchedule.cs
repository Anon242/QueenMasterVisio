using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Visio.Page;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Reflection.Emit;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.ObjectModel;

namespace QueenMasterVisio
{
    // Класс для собирания информации с документа для импорта в csv 
    internal class CableSchedule
    {
        /*

Cable designation;Cable route;;;Cable;;;;
;From;Way;To;Type;Size;Length;Voltage;Note
Ax;Shield;A15.3;G15.3;ШВВП;3 x 1.5mm2;;230v;Ok
Ax;Shield;A15.3;G15.3;ШВВП;3 x 1.5mm2;;230v;Ok

*/
        // flexible cord
        /*
         * Одно из девайсов может быть свет, надо определить вольтаж
         * Кстати есть метод что бы свапнуть трасер если что
         * 
         */

        // Списки которые мы обрабатываем на from и to, например в To не может быть Shield или From Camera тоже быть не может
        // Распред коробки
        readonly static string[] shapeTypeFrom = new string[] {"Device", "Light", "Shield", "Recorder", "Box"};
        readonly static string[] shapeTypeTo = new string[] { "Device", "Light", "Alarm","Sound","Camera", "Box"};

        public static string Generate(Visio.Page page)
        {
            // Проверка на дурака
            if (!Tools.CellExistsCheck(page, "User.pageCode"))
                return null;
            if (Tools.CellFormulaGet(page, "User.pageCode") != "Plan")
                return null;

            string result = "Designation;From;Way;To;Type;Voltage;Length;Note\r\n";

            foreach (Visio.Shape shape in page.Shapes)
            {
                if(Checker.isLine(shape))
                {
                    if (shape.Connects.Count >= 2)
                    {
                        Visio.Shape connectedShapeFrom = shape.Connects[1].ToSheet;
                        Visio.Shape connectedShapeTo = shape.Connects[2].ToSheet;

                        if (shapeTypeFrom.Any(type => connectedShapeFrom.Name.Contains(type)) && shapeTypeTo.Any(type => connectedShapeTo.Name.Contains(type)))
                        {
                            string nameValueFrom = connectedShapeFrom.CellsU["Prop.Number"].FormulaU.Replace("\"", "");
                            string nameValueTo = connectedShapeTo.CellsU["Prop.Number"].FormulaU.Replace("\"", "");
                            Wire wire = WireService.GetWireByColorName(shape.CellsU["LineColor"].FormulaU);
                            if (wire == null) { continue; }
                            // Ax;Shield;A15.3;G15.3;ШВВП;3 x 1.5mm2;;230v;Ok

                            // Пробуем смотреть есть ли у него Prop.Number и Prop.Type
                            string propType = Tools.CellFormulaGet(shape, "Prop.Type");

                            // type
                            string f1 = wire.name + " - " + wire.comment;
                            // from
                            string f2 = (wire.name == "Lx" ? "L" : "G") + nameValueFrom;
                            if (connectedShapeFrom.Name.Contains("Shield"))
                                f2 = "G" + nameValueFrom + " - Shield";
                            else if (connectedShapeFrom.Name.Contains("Recorder"))
                                f2 = "G" + nameValueFrom + " - Recorder";
                            else if (connectedShapeFrom.Name.Contains("Box"))
                                f2 = "J" + nameValueFrom;

                            // way
                            string f3 = wire.name[0] + nameValueTo;
                            if (connectedShapeTo.Name.Contains("Box"))
                                f3 = "PJ" + nameValueTo;
                            // to
                            string f4 = (wire.name == "Lx" ? "L" : "G") + nameValueTo;
                            if(wire.name == "Yx")
                                f4 = "Y" + nameValueTo;
                            else if (wire.name == "Ax")
                                f4 = "A" + nameValueTo;
                            else if (wire.name == "Vx")
                                f4 = "V" + nameValueTo;
                            else if (connectedShapeTo.Name.Contains("Box"))
                                f4 = "J" + nameValueTo;
                            // sech
                            string f6 = wire.defaultCable == null ? "" : wire.defaultCable;
                            if (wire.name == "Lx" && Tools.CellExistsCheck(connectedShapeTo, "Prop.Type") && connectedShapeTo.CellsU["Prop.Type"].ResultStrU[""] == "RGB")
                            {
                                f6 = "4 x 1.5";
                            }
                            else if (propType != null && propType != "")
                            {
                                f6 = propType;
                            }
                            // volt
                            string f7 = wire.voltage;
                            if (wire.name == "Lx" && Tools.CellExistsCheck(connectedShapeTo, "Prop.Type"))
                            {
                                f7 = connectedShapeTo.CellsU["Prop.Type"].ResultStrU[""] == "RGB" ? "12v" : connectedShapeTo.CellsU["Prop.Type"].ResultStrU[""];
                            }
                            //len null
                            string f8 = "";
                            // note null
                            string f9 = "";

                            result += f1 + ";" + f2 + ";" + f3 + ";" + f4 + ";" + f6 + ";" + f7 + ";" + f8 + ";" + f9 + "\r\n";
                            // Если у нас Rx, то там 2 кабеля, делаем дубликат сразу по питанию
                            if(wire.name == "Rx")
                            {
                                f6 = " 3 x 1.5";
                                result += f1 + ";" + f2 + ";" + f3 + ";" + f4 + ";" + f6 + ";" + f7 + ";" + f8 + ";" + f9 + "\r\n";
                            }
                        }
                    }
                }
            }

            return result;
        }

    }
}
