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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpened += new Visio.EApplication_DocumentOpenedEventHandler(OnDocumentOpened);
            this.Application.BeforeDocumentClose += new Visio.EApplication_BeforeDocumentCloseEventHandler(OnBeforeDocumentClose);
            this.Application.PageAdded += new Visio.EApplication_PageAddedEventHandler(OnPageAdded);
            this.Application.PageChanged += new Visio.EApplication_PageChangedEventHandler(OnPageChanged);

        }

        private void OnPageChanged(Page Page)
        {
            pageExplorer.UpdateExplorer();
        }
        private void OnPageAdded(Page Page)
        {
            pageExplorer.UpdateExplorer();
        }

        private void CreateEmbeddedWindow()
        {
            try
            {
                // Создаем встроенное окно в Visio
                customWindow = this.Application.ActiveWindow.Windows.Add("Проводник по документам",                                           // Заголовок
                    (Visio.VisWindowStates.visWSVisible |
                           Visio.VisWindowStates.visWSDockedRight),          // Состояние - видимое, закреплено справа
                    (short)Visio.VisWinTypes.visAnchorBarAddon,              // Тип - панель дополнения
                    0, 0,                                                    // Позиция
                    300, 600,                                                // Размер
                    "", "", 0                                                // Параметры слияния
                );
                // Создаем UserControl
                pageExplorer = new Explorer(this.Application, customWindow);

                // Встраиваем UserControl в окно Visio
                EmbedUserControlInWindow();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating embedded window: {ex.Message}");
            }
        }

        private void CreateEmbeddedWindowChangeLog(string path)
        {
            this.Application.OnComponentEnterState(Visio.VisOnComponentEnterCodes.visComponentStateModal,true);
            try
            {
                // Создаем встроенное окно в Visio
                customWindowChangeLog = this.Application.ActiveWindow.Windows.Add("Change Log",   // Заголовок
                    (Visio.VisWindowStates.visWSVisible |
                           Visio.VisWindowStates.visWSFloating),          // Состояние - видимое, закреплено справа
                    (short)Visio.VisWinTypes.visAnchorBarAddon,              // Тип - панель дополнения
                    200, 200,                                                    // Позиция
                    348, 220,                                               // Размер
                    "", "", 0                                                // Параметры слияния
                );

                // Создаем UserControl
                changeLog = new ChangeLog.ChangeLog(customWindowChangeLog,this.Application, path);
                // Change log
                EmbedUserControlInWindowChangeLog();
                
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating embedded window: {ex.Message}");
            }
            finally
            {
                this.Application.OnComponentEnterState(Visio.VisOnComponentEnterCodes.visComponentStateModal,false);
            }
        }
        private void EmbedUserControlInWindow()
        {
            try
            {
                // Получаем handle окна Visio
                IntPtr windowHandle = new IntPtr(customWindow.WindowHandle32);

                // Получаем handle UserControl
                pageExplorer.CreateControl();
                IntPtr controlHandle = pageExplorer.Handle;

                // Устанавливаем UserControl как дочернее окно
                SetParent(controlHandle, windowHandle);

                // Устанавливаем стили окна
                SetWindowLong(controlHandle, GWL_STYLE, WS_CHILD | WS_VISIBLE);

                // Растягиваем UserControl на все окно
                pageExplorer.Dock = DockStyle.Fill;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error embedding control: {ex.Message}");
            }
        }

        private void EmbedUserControlInWindowChangeLog()
        {
            try
            {
                // Получаем handle окна Visio
                IntPtr windowHandle = new IntPtr(customWindowChangeLog.WindowHandle32);

                // Получаем handle UserControl
                changeLog.CreateControl();
                IntPtr controlHandle = changeLog.Handle;

                // Устанавливаем UserControl как дочернее окно
                SetParent(controlHandle, windowHandle);

                // Устанавливаем стили окна
                SetWindowLong(controlHandle, GWL_STYLE, WS_CHILD | WS_VISIBLE);

                // Растягиваем UserControl на все окно
                changeLog.Dock = DockStyle.Fill;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error embedding control: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private void OnBeforeDocumentClose(Document doc)
        {
            //myPage.closeThread();
        }
        private void OnDocumentOpened(Document doc)
        {
            // Это скрипт на шаблонs открылся
            if (doc.Name.Contains(".vss"))
                return;
            // Explorer
            CreateEmbeddedWindow();
            Debug.WriteLine("Открыт документ и развернут explorer");


            //this.Application.BeforeDocumentSave += new Visio.EApplication_BeforeDocumentSaveEventHandler(Application_BeforeDocumentSave);
            this.Application.DocumentSaved += new Visio.EApplication_DocumentSavedEventHandler(Application_DocumentSaved);
            myPage = new VisioEventAggregator(this.Application, pageExplorer);
            this.Application.ShapeChanged += new Visio.EApplication_ShapeChangedEventHandler(myPage.OnShapeChanged);
            this.Application.ShapeAdded += new Visio.EApplication_ShapeAddedEventHandler(myPage.OnShapeAdded);
            myPage.start();

            pageExplorer.UpdateExplorer();

            links = new LocalLinks(doc.FullName);

            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 3000;
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    

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

                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error preparing changelog: {ex.Message}");
                }
            };
            timer.Start();
        }

        private DateTime RoundToMinute(DateTime dt)
        {
            return new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, 0);
        }

        private void Application_DocumentSaved(Visio.Document doc)
        {
            try
            {
                // Заголовки
                string headers = pageExplorer.GetHeadlinesText();
                if (!string.IsNullOrEmpty(headers))
                {
                    string headersPath = System.IO.Path.Combine(links.MetafilesDir, "headers.txt");
                    File.WriteAllText(headersPath, headers);
                }

                // Дальше Чейнджлог
                CreateEmbeddedWindowChangeLog(links.ChangeLogsDir);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error preparing changelog: {ex.Message}");
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            myRibbonTracer = new MainLentXml();
            return myRibbonTracer;
          
        }




        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
