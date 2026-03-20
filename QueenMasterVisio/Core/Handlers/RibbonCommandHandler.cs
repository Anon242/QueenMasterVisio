using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Managers;
using QueenMasterVisio.Core.Services;
using QueenMasterVisio;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

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

        
        public void onRibbonTracerBtnDevice(IRibbonControl control)
        {
            var visioApp = Globals.ThisAddIn.Application;
            var activePage = visioApp.ActivePage;
            if (control.Id.Contains("btn"))
                onButtonPressedDevice(activePage, control.Id);

        }
        

        private void onButtonPressedPlan(Page page, string buttonId)
        {
            if (!page.IsPlanPage())
                return;


            switch (buttonId)
            {
                case "btnAutoConnect":
                    if (page.GetPlanCode() != "Plan" && page.GetPlanCode() != "All")
                        WireAutoConnectionService.autoConnect(page);
                    break;
                case "btnReload":
                    MasterAutoPlanPageService.CreateNewReloadPages(page);
                    break;
                ////////////////////////////////// Слои
                case "btnAll":
                case "btnPlan":
                case "btnOther_1":
                case "btnOther_2":
                case "btnPx":
                case "btnEx":
                case "btnRx":
                case "btnDx":
                case "btnCx":
                case "btnSx":
                case "btnYx":
                case "btnVx":
                case "btnLx":
                case "btnAx":
                    // Если будет проблема, тогда будем хранить план код в User.Shape
                    PageManager.SetOptionsAllPlanlayer(page, buttonId);
                    break;
                ////////////////////////////////// Слои
                case "btnGetLineData":
                    Clipboard.SetText(CableService.Generate(page));
                    break;
                case "btnSetHyperLinks":
                    //SetHyperLinks(page);
                    break;
                case "btnLookDevices":
                    //explorer.LookDevices(page);
                    break;
                case "btnDevicesCheck":
                    if (!page.IsPlanPage())
                        return;
                    CheckerService.CheckDevicesInPlan(page); // Не верно 
                    break;
            }
        }

        private void onButtonPressedDevice(Page page, string buttonId)
        {

            switch (buttonId)
            {
                case "btnLock":
                    if (page.IsPlanPage())
                        return;

                    PageManager.LockOrUnlockLayer(page);
                    break;
                case "btnUpdatePage":
                    if (!page.IsAutoTracePage())
                        return;
                    
                    PageManager.RedrawPageAuto(page);
                    break;

                // Сброс линий 
                case "btnResetLine":
                    CableService.ResetLines(page);
                    break;
                case "btnCreatePlan":
                    PageManager.CreateNewPlan(page);
                    
                    break;
                case "btnCopyAll":
                    try
                    {
                        PageManager.LockAllLayers(page,false);
                        Selection selection = page.Application.ActiveWindow.Selection;
                        selection.SelectAll();
                        selection.Copy(VisCutCopyPasteCodes.visCopyPasteNoTranslate | VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
                        selection.DeselectAll();

                    }
                    catch{}
                    break;

                case "btnPasteAll": 
                    try
                    {
                        using (VisioEventSuppressor.SuppressShapeAdded())
                        {
                            // Страница заблокирована
                            if (RedSquareCreator.RedSquareGetLayer(page) != null)
                                return;
                            // Нет наших данных в буфере обмена
                            if (!(Clipboard.ContainsData("Visio 11.0 Shapes") || Clipboard.ContainsData("Visio 15.0 Shapes") || Clipboard.ContainsData("Visio 15.0 Text")))
                                return;

                            page.Paste(VisCutCopyPasteCodes.visCopyPasteNoTranslate | VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
                            
                            Clipboard.Clear();
                            // На случай если мы копировали с заблоканого слоя
                            RedSquareCreator.RedSquareDelete(page);
                        }
                    }
                    catch{}

                    break;
                case "btnLookDevicesOnPlan":
                    //DeviceCheck.DeviceCheck deviceCheck = new DeviceCheck.DeviceCheck(page.Application);
                    //deviceCheck.Show();
                    Form1 form = new Form1();
                    form.Show();

                    break;
                case "btnCreateNewDevice":
                    short pageIndex = (short)(page.Document.Pages.Count - 1);

                    Page newPage = DocumentManager.CreateNewPage(VisioEventAggregator.explorer.ShowRenameDialog("G" + pageIndex));
                    
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
                    break;
            }
        }
    }
}
