using Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using Page = Microsoft.Office.Interop.Visio.Page;
using Section = Microsoft.Office.Interop.Visio.Section;
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
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Visio;
using System.Windows.Shapes;
using Shape = Microsoft.Office.Interop.Visio.Shape;

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
                    try
                    {
                        scopeId = Globals.ThisAddIn.Application.BeginUndoScope("Создание автостраниц");
                        MasterAutoPlanPageService.CreateNewReloadPages(page);
                    }
                    catch (System.Exception)
                    {
                        throw;
                    }
                    finally
                    {
                        Globals.ThisAddIn.Application.EndUndoScope(scopeId, true);
                        page.Document.UndoEnabled = true;
                    }
   
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

                case "btnCreateNewDeviceInPlan":
                    Microsoft.Office.Interop.Visio.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    CreateDeviceInPlan.CreatePage(page, selection);
                    break;
            }
        }

        private void onButtonPressedDevice(Page page, string buttonId)
        {

            switch (buttonId)
            {
                case "btnLock":
                    if (page.IsPlanPage() || page.IsAutoTracePage())
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
                    //CableService.ResetLines(page);
                    if (page.IsPlanPage() || page.IsAutoTracePage())
                        break;

                    System.Windows.Window myWindow1 = new System.Windows.Window();
                    myWindow1.Title = "Binding";
                    myWindow1.Content = new Binding.BindingForm(page, myWindow1);
                    myWindow1.MinWidth = 800;
                    myWindow1.MinHeight = 450;
                    myWindow1.Width = 800;
                    myWindow1.Height = 450;
                    myWindow1.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    myWindow1.Topmost = true;

                    myWindow1.Show();

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
                    if (page.IsPlanPage() || page.IsAutoTracePage())
                    {
                        System.Windows.Forms.MessageBox.Show("По техническим причинам, создать новый девайс находясь на плане не возможно, перейдите на другую страницу с девайсом и попробуйте снова",
                                "Ошибка",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                        break;
                    }

                    System.Windows.Window myWindow2 = new System.Windows.Window();

                    myWindow2.Title = "Binding";
                    myWindow2.Content = new Binding.BindingFormCreator(page, myWindow2);
                    myWindow2.MinWidth = 800;
                    myWindow2.MinHeight = 450;
                    myWindow2.Width = 800;
                    myWindow2.Height = 450;
                    myWindow2.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    myWindow2.Topmost = true;
                            
                    myWindow2.Show();


                    break;
                case "btnGetLines":
                    if (!(page.GetUserPageCode() == "Device" || page.GetUserPageCode() == "Light"))
                    {
                        System.Windows.Forms.MessageBox.Show("Страница не является девайсом или светом", "Ошибка", MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        break;
                    }
                    string deviceList = page.GetCellFormulaU("User.deviceList");
                    if (string.IsNullOrEmpty(deviceList))
                    {
                        System.Windows.Forms.MessageBox.Show("Страница не привязана", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                    }
                    
                    // Надо только один раз открыть!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        Microsoft.Office.Interop.Visio.Document sourceStencil = Globals.ThisAddIn.Application.Documents.OpenEx(ThisAddIn.links.QueenFigures,
        (short)(Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenHidden | Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDontList));
                  
                    Microsoft.Office.Interop.Visio.Master masterCable = sourceStencil.Masters.get_ItemU("Cable");
                    Microsoft.Office.Interop.Visio.Master masterABCD = sourceStencil.Masters.get_ItemU("ABCD");

                    List<Shape> shapes = new List<Shape>();

                    // Получить с странциы все девайсо страницы
                    // Пройтись по каждому, получить их да да нет нет px ax и тд

                    foreach (string item in deviceList.Split(';'))
                    {
                        string[] deviceSplit = item.Split('-');
                        string num = deviceSplit[0].Trim();
                        string deviceNameU = deviceSplit[1].Trim();
                        string pageId = deviceSplit[2].Trim();

                        shapes.Add(page.Document.Pages.ItemFromID[int.Parse(pageId)].Shapes[deviceNameU]);
                    }
                   
                    foreach (Shape shape in shapes)
                    {
                        Debug.WriteLine(shape.GetCellResultString("Prop.Px"));
                    }

                    
                    
                    break;
            }
        }
    }
}
