using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Handlers;
using QueenMasterVisio.Core.Services;
using QueenMasterVisio.Ribbon;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;


namespace QueenMasterVisio
{
    public partial class ThisAddIn
    {

        public VisioEventAggregator myPage;
        MainLentXml myRibbonTracer;

        private Explorer pageExplorer;
        private ChangeLog.ChangeLog changeLog;

        private Visio.Window customWindow;
        private Visio.Window customWindowChangeLog;

        public static LocalLinks links;

        [DllImport("user32.dll")]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_STYLE = -16;
        private const int WS_CHILD = 0x40000000;
        private const int WS_VISIBLE = 0x10000000;
        private void SafeExecute(Action action, string context = null)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                string message = $"Произошла критическая ошибка в надстройке QueenMasterVisio{(string.IsNullOrEmpty(context) ? "" : $" ({context})")}:\n\n{ex.Message}\n\n{ex.StackTrace}";
                MessageBox.Show(message, "Ошибка надстройки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine(message);
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SafeExecute(() =>
            {
                this.Application.DocumentOpened += new Visio.EApplication_DocumentOpenedEventHandler(OnDocumentOpened);
                this.Application.BeforeDocumentClose += new Visio.EApplication_BeforeDocumentCloseEventHandler(OnBeforeDocumentClose);
                this.Application.PageAdded += new Visio.EApplication_PageAddedEventHandler(OnPageAdded);
                this.Application.PageChanged += new Visio.EApplication_PageChangedEventHandler(OnPageChanged);
            }, "Startup");
        }

        private void OnPageChanged(Page Page)
        {
            SafeExecute(() => pageExplorer?.UpdateExplorer(), "OnPageChanged");
        }

        private void OnPageAdded(Page Page)
        {
            SafeExecute(() => pageExplorer?.UpdateExplorer(), "OnPageAdded");
        }

        private void CreateEmbeddedWindow()
        {
            SafeExecute(() =>
            {
                customWindow = this.Application.ActiveWindow.Windows.Add("Проводник по документам",
                    (Visio.VisWindowStates.visWSVisible | Visio.VisWindowStates.visWSDockedRight),
                    (short)Visio.VisWinTypes.visAnchorBarAddon,
                    0, 0,
                    300, 600,
                    "", "", 0
                );
                pageExplorer = new Explorer(this.Application, customWindow);
                EmbedUserControlInWindow();
            }, "CreateEmbeddedWindow");
        }

        private void CreateEmbeddedWindowChangeLog(string path)
        {
            SafeExecute(() =>
            {
                this.Application.OnComponentEnterState(Visio.VisOnComponentEnterCodes.visComponentStateModal, true);
                try
                {
                    customWindowChangeLog = this.Application.ActiveWindow.Windows.Add("Change Log",
                        (Visio.VisWindowStates.visWSVisible | Visio.VisWindowStates.visWSFloating),
                        (short)Visio.VisWinTypes.visAnchorBarAddon,
                        200, 200,
                        348, 220,
                        "", "", 0
                    );
                    changeLog = new ChangeLog.ChangeLog(customWindowChangeLog, this.Application, path);
                    EmbedUserControlInWindowChangeLog();
                }
                finally
                {
                    this.Application.OnComponentEnterState(Visio.VisOnComponentEnterCodes.visComponentStateModal, false);
                }
            }, "CreateEmbeddedWindowChangeLog");
        }

        private void EmbedUserControlInWindow()
        {
            SafeExecute(() =>
            {
                IntPtr windowHandle = new IntPtr(customWindow.WindowHandle32);
                pageExplorer.CreateControl();
                IntPtr controlHandle = pageExplorer.Handle;
                SetParent(controlHandle, windowHandle);
                SetWindowLong(controlHandle, GWL_STYLE, WS_CHILD | WS_VISIBLE);
                pageExplorer.Dock = DockStyle.Fill;
            }, "EmbedUserControlInWindow");
        }

        private void EmbedUserControlInWindowChangeLog()
        {
            SafeExecute(() =>
            {
                IntPtr windowHandle = new IntPtr(customWindowChangeLog.WindowHandle32);
                changeLog.CreateControl();
                IntPtr controlHandle = changeLog.Handle;
                SetParent(controlHandle, windowHandle);
                SetWindowLong(controlHandle, GWL_STYLE, WS_CHILD | WS_VISIBLE);
                changeLog.Dock = DockStyle.Fill;
            }, "EmbedUserControlInWindowChangeLog");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            SafeExecute(() =>
            {
                // Здесь можно добавить код для корректного завершения работы надстройки
            }, "Shutdown");
        }

        private void OnBeforeDocumentClose(Document doc)
        {
            SafeExecute(() =>
            {
                //myPage.closeThread();
            }, "OnBeforeDocumentClose");
        }

        private void OnDocumentOpened(Document doc)
        {
            if (!doc.Name.Contains(".vsdx"))
                return;

            SafeExecute(() =>
            {

                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 6000;
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    SafeExecute(() =>
                    {

                        CreateEmbeddedWindow();
                        Debug.WriteLine("Открыт документ и развернут explorer");

                        this.Application.DocumentSaved += new Visio.EApplication_DocumentSavedEventHandler(Application_DocumentSaved);
                        myPage = new VisioEventAggregator(this.Application, pageExplorer);
                        this.Application.ShapeChanged += new Visio.EApplication_ShapeChangedEventHandler(myPage.OnShapeChanged);
                        this.Application.ShapeAdded += new Visio.EApplication_ShapeAddedEventHandler(myPage.OnShapeAdded);
                        myPage.start();

                        pageExplorer.UpdateExplorer();

                        links = new LocalLinks(doc.FullName);

                        if (links.isLocalFile)
                        {
                            MessageBox.Show("Документ локальный, запрет вывода окна сохранения и генерации хедеров (временное предупреждение)",
                                "Локальный документ",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                            return;
                        }

                        Debug.WriteLine("Получили чейгнджлоги");

                        var lastFile = new DirectoryInfo(links.ChangeLogsDir).GetFiles().OrderByDescending(f => f.CreationTime).FirstOrDefault();
                        Debug.WriteLine(lastFile);
                        if (lastFile != null)
                        {
                            string fileName = lastFile.Name;
                            DateTime creationDate = RoundToMinute(lastFile.LastWriteTime);
                            DateTime vsdxLast = RoundToMinute(File.GetLastWriteTime(links.BasePath));

                            Debug.WriteLine(creationDate);
                            Debug.WriteLine(vsdxLast);

                            if (Math.Abs((creationDate - vsdxLast).TotalMinutes) > 1)
                            {
                                MessageBox.Show("Возможно вы используете старую (локальную) версию этого документа, проверьте что onedrive включен и попробуйте использовать \"Освободить место\" или же, если вы уверены что документ актуальный, проигнорируйте это сообщение",
                                    "Ой",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                            }
                        }
                    }, "Timer_Tick");
                };
                timer.Start();
            }, "OnDocumentOpened");
        }

        private DateTime RoundToMinute(DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, 0);
        }

        private void Application_DocumentSaved(Visio.Document doc)
        {
            SafeExecute(() =>
            {
                if (links.isLocalFile)
                {
                    MessageBox.Show("Документ локальный, запрет вывода окна сохранения и генерации хедеров (временное предупреждение)",
                        "Локальный документ",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                string headers = pageExplorer.GetHeadlinesText();
                if (!string.IsNullOrEmpty(headers))
                {
                    string headersPath = System.IO.Path.Combine(links.MetafilesDir, "headers.txt");
                    File.WriteAllText(headersPath, headers);
                }
                CreateEmbeddedWindowChangeLog(links.ChangeLogsDir);
            }, "Application_DocumentSaved");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                myRibbonTracer = new MainLentXml();
                return myRibbonTracer;
            }
            catch (Exception ex)
            {
                string message = $"Ошибка инициализации ленты: {ex.Message}\n\n{ex.StackTrace}";
                MessageBox.Show(message, "Ошибка надстройки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine(message);
                return null;
            }
        }

        #region Код, автоматически созданный VSTO

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
