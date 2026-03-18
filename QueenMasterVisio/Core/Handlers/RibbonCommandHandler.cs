using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueenMasterVisio.Core.Handlers
{
    internal class RibbonCommandHandler
    {
        public void onRibbonTracerBtnPlan(IRibbonControl control)
        {
            var visioApp = Globals.ThisAddIn.Application;
            var activePage = visioApp.ActivePage;
            if (control.Id.Contains("btn"))
                onButtonPressedPlan(activePage, control.Id);

        }

        /*
        public void onRibbonTracerBtnDevice(IRibbonControl control)
        {
            var visioApp = Globals.ThisAddIn.Application;
            var activePage = visioApp.ActivePage;
            if (control.Id.Contains("btn"))
                onButtonPressedDevice(activePage, control.Id.Substring(3));

        }
        */

        private void onButtonPressedPlan(Page page, string buttonId)
        {
            if (!page.IsPlanPage())
                return;

            string activePlanCode = page.GetPlanCode();

            switch (buttonId)
            {
                case "AutoConnect":
                    if (activePlanCode != "Plan" && activePlanCode != "All")
                        Services.WireAutoConnectionService.autoConnect(page);
                    break;
                default:
                    break;
            }

            if (buttonId == "AutoConnect")
            {

            }
            else if (buttonId == "Reload")
            {
                onReloadBtnPressed(page);
            }
            else if (buttonId == "CreatePlan")
            {
                onCreatePlanPressed(page);
            }
            // Проверка по второму символу, осторожно
            else if (buttonId[1] == 'x' || buttonId == "Plan" || buttonId == "All" || buttonId.Contains("Other"))
            {
                activePlanCode = buttonId;
                onLayersBtnPressed(page);
            }
            else if (buttonId == "Lock" && GetActivePageCode() != "Plan")
            {
                foreach (Visio.Shape shape in page.Shapes)
                {
                    if (CheckerService.isLine(shape))
                    {
                        Tools.CellFormulaSet(shape, "ConFixedCode", "2");
                    }
                }
                setRedSquareOnPage(page, true);
                lockAllLayers(page);
            }
            else if (buttonId == "Unlock" && GetActivePageCode() != "Plan")
            {
                unlockAllLayers(page);
                setRedSquareOnPage(page, false);
            }
            else if (buttonId == "UpdatePage" && GetActivePageCode() == "planAuto")
            {
                redrawPageAuto(page);
            }
            else if (buttonId == "DevicesCheck" && GetActivePageCode() == "Plan")
            {
                CheckerService.CheckDevicesInPlan(page); // Не верно 
            }
            else if (buttonId == "CopyAll")
            {
                try
                {
                    // ЭТО ВРЕМЯНКА, КОСТЫЛЬ БЛЯТЬ
                    foreach (Visio.Shape shape in page.Shapes)
                    {
                        if (CheckerService.isLine(shape))
                        {
                            Tools.CellFormulaSet(shape, "ConFixedCode", "2");
                        }
                    }
                    unlockAllLayers(page);
                    Selection selection = page.Application.ActiveWindow.Selection;
                    selection.SelectAll();
                    selection.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | Visio.VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
                    lockAllLayers(page);
                    selection.DeselectAll();
                }
                catch
                {


                }
            }
            else if (buttonId == "PasteAll")
            {

                try
                {
                    onShapeAddedBreak = true;
                    page.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | Visio.VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
                    setRedSquareOnPage(page, false);
                    Clipboard.Clear();
                    Selection selection = page.Application.ActiveWindow.Selection;
                    selection.SelectAll();
                    selection.ConvertToGroup();
                    onShapeAddedBreak = false;
                }
                catch
                {


                }
                /*
				if (masterLineCounts == 0) {
					bool buffer = true;
					while (buffer) {
						buffer = false;
                        foreach (Master master in page.Document.Masters)
							if (master.Name.Contains("Динамическая соединительная линия."))
							{
								buffer = true;
                                master.Delete();
							}
					}
				}
				*/
            }
            // В буфер обмена вставляем строки 
            else if (buttonId == "GetLineData")
            {
                //MessageBox.Show(CableSchedule.Generate(page),"Test",MessageBoxButtons.OK);
                string str = CableService.Generate(page);
                Clipboard.SetText(str);
                Debug.WriteLine(str);

            }
            else if (buttonId == "SetHyperLinks")
            {
                SetHyperLinks(page);
            }
            else if (buttonId == "LookDevices")
            {
                explorer.LookDevices(page);
            }
            else if (buttonId == "LookDevicesOnPlan")
            {
                // Получаем соединения с плана находясь на странице устройства 
                LookDevicesOnPlan(page);
            }
            else if (buttonId == "CreateNewDevice")
            {
                // Создать новый девайс
                Page activePage = page;
                Page newPage = page.Document.Pages.Add();
                app.ActiveWindow.Page = activePage;
                short pageIndex = (short)(page.Document.Pages.Count - 1);
                newPage.Name = explorer.ShowRenameDialog("G" + pageIndex);

                if (newPage.Name[0] == 'G')
                {
                    // Делаем что она была перед первым светом
                    foreach (Page _page in page.Document.Pages)
                    {
                        Regex regexLight = new Regex(@"^L\d");
                        if (regexLight.IsMatch(_page.Name))
                        {
                            newPage.Index = _page.Index;
                            break;
                        }
                    }
                }
                else
                {
                    newPage.Index = (short)(pageIndex);
                }

                app.ActiveWindow.Page = newPage;
            }
        }

    }
}
