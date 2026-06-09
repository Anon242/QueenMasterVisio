using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;

namespace QueenMasterVisio.Core.Services
{
    internal class CableAutoDeviceService
    {
        static public void DropAutoCableInLine(Page page, Shape shape, Shape line, float x, float y, bool isRx = false)
        {

            Microsoft.Office.Interop.Visio.Master masterCable = Globals.ThisAddIn.sourceStencil.Masters.get_ItemU("Cable");

            bool isComing = false;
            string _num = line.GetCellFormulaU("User.To").Substring(1);
            string _numDevice = shape.GetCellFormulaU("Prop.number");
            if (_num == _numDevice)
                isComing = true;

            Shape newCable = page.Drop(masterCable, x, y);

            newCable.SetUserCell("BindingLine", line.NameU);

            newCable.SetFormula("Prop.Name", $"=IFERROR(Pages[{shape.ContainingPage.NameU}]!Sheet.{line.ID}!User.Way, \"NIL\")", true);

            if (isComing)
                newCable.SetFormula("Prop.Device", $"=IFERROR(Pages[{shape.ContainingPage.NameU}]!Sheet.{line.ID}!User.From,\"NIL\")", true);
            else
                newCable.SetFormula("Prop.Device", $"=IFERROR(Pages[{shape.ContainingPage.NameU}]!Sheet.{line.ID}!User.To,\"NIL\")", true);

            if(!isRx)
                newCable.SetFormula("Prop.type", $"=IFERROR(SUBSTITUTE(Pages[{shape.ContainingPage.NameU}]!Sheet.{line.ID}!User.OverCable,\" Cat 5E\",\"\",1),\"NIL\")", true);
            else
                newCable.SetFormula("Prop.type", $"3x1.5");


            newCable.SetFormula("Prop.Name.Invisible", "1");
            newCable.SetFormula("Prop.Device.Invisible", "1");
            newCable.SetFormula("Prop.type.Invisible", "1");
        }

        
        static public void DropAutoABCDInLine(Page page, Shape shape, Shape line, float x, float y)
        {
            Microsoft.Office.Interop.Visio.Master masterABCD = Globals.ThisAddIn.sourceStencil.Masters.get_ItemU("ABCD");
            Shape ABCD;
            ABCD = page.Drop(masterABCD, x, y);
        } 
    }
}
