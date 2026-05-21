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
using QueenMasterVisio.DeviceControl;
using System.Windows;
using System.Collections.Generic;

namespace QueenMasterVisio.Core.Handlers
{
    internal class RibbonCommandHandler
    {
        System.Windows.Window window;
        public void onRibbonTracerBtnPlan(IRibbonControl control)
        {
            var visioApp = Globals.ThisAddIn.Application;
            var activePage = visioApp.ActivePage;
            if (control.Id.Contains("btn"))
                onButtonPressedPlan(activePage, control.Id);

        }

        public void CreateFormDeviceControl()
        {
            
            window = new System.Windows.Window
            {
                Title = "Device Control",
                Content = new DeviceControlForm(Globals.ThisAddIn.Application),
                MinWidth = 800,
                MinHeight = 450,
                Width = 800,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Topmost = true
            };
            window.Show();
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
            int scopeId = 0;

            switch (buttonId)
            {
                case "btnAutoConnect":
                    if (page.GetPlanCode() != "Plan" && page.GetPlanCode() != "All")
                        WireAutoConnectionService.autoConnect(page);
                    break;
                case "btnReload":
                    scopeId = Globals.ThisAddIn.Application.BeginUndoScope("Создание автостраниц");
                    MasterAutoPlanPageService.CreateNewReloadPages(page);
                    Globals.ThisAddIn.Application.EndUndoScope(scopeId, true);

                    page.Document.UndoEnabled = true;
                    break;
                ////////////////////////////////// Слои
                case "btnAll":
                case "btnPlan":
                case "btnOther1":
                case "btnOther2":
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
                    scopeId = Globals.ThisAddIn.Application.BeginUndoScope("Переключение кнопок плана");
                    // Если будет проблема, тогда будем хранить план код в User.Shape
                    PageManager.SetOptionsAllPlanlayer(page, buttonId);
                    Globals.ThisAddIn.Application.EndUndoScope(scopeId, true);

                    break;
                ////////////////////////////////// Слои
                case "btnGetLineData":
                    System.Windows.Clipboard.SetText(CableService.Generate(page));
                    break;
                case "btnSetHyperLinks":
                    //SetHyperLinks(page);
                    break;
           
                case "btnDevicesCheck":
                    if (!page.IsPlanPage())
                        return;


                    CreateFormDeviceControl();
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
                case "btnLookLogs":
                    try
                    {
                        List<string> paths = new List<string>();
                        string changeLogsDir = ThisAddIn.links.ChangeLogsDir;
                        foreach (var file in System.IO.Directory.GetFiles(changeLogsDir))
                        {
                            if (file.Contains("Changelog_"))
                                paths.Add(file);
                        }
                       
                        window = new System.Windows.Window
                        {
                            Title = "Logs",
                            Content = new LogViewer.LogView(paths),
                            MinWidth = 800,
                            MinHeight = 450,
                            Width = 800,
                            Height = 450,
                            WindowStartupLocation = WindowStartupLocation.CenterScreen,
                            Topmost = true
                        };
                        window.Show();
                    }
                    catch (System.Exception)
                    {

                        throw;
                    }
            
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
                            if (!(System.Windows.Clipboard.ContainsData("Visio 11.0 Shapes") || System.Windows.Clipboard.ContainsData("Visio 15.0 Shapes") || System.Windows.Clipboard.ContainsData("Visio 15.0 Text")))
                                return;

                            page.Paste(VisCutCopyPasteCodes.visCopyPasteNoTranslate | VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
                            
                            System.Windows.Clipboard.Clear();
                            // На случай если мы копировали с заблоканого слоя
                            RedSquareCreator.RedSquareDelete(page);
                        }
                    }
                    catch{}

                    break;
                case "btnLookDevicesOnPlan":
                    CreateFormDeviceControl();

                    break;
                case "btnCreateNewDevice":
                    if (page.IsPlanPage())
                    {
                        System.Windows.Forms.MessageBox.Show("По техническим причинам, создать новый девайс находясь на плане не возможно, перейдите на другую страницу с девайсом и попробуйте снова",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        return;
                    }
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
