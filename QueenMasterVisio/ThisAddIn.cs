using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;


namespace QueenMasterVisio
{
    public partial class ThisAddIn
    {
        MyPage myPage;
        MyRibbonTracer myRibbonTracer;

        private Explorer pageExplorer;
        private Visio.Window customWindow;

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

            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                Exception ex = args.ExceptionObject as Exception;
                string message = ex?.Message ?? "Неизвестная ошибка";
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };

            // Обработчик для UI потоков
            System.Windows.Forms.Application.ThreadException += (s, args) =>
            {
                Exception ex = args.Exception as Exception;
                string message = ex?.Message ?? "Неизвестная ошибка";
                MessageBox.Show(message + "\n UI", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
        }

        private void OnPageChanged(Page Page)
        {
            pageExplorer.UpdateExplorer(Page.Application.ActiveDocument);
        }
        private void OnPageAdded(Page Page)
        {
            pageExplorer.UpdateExplorer(Page.Application.ActiveDocument);
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
            if (doc.Name.Contains(".vssx"))
                return;
            // Explorer
            // Создаем окно только если customWindow еще не существует или был удален
            if (customWindow == null || pageExplorer == null || pageExplorer.IsDisposed)
            {
                CreateEmbeddedWindow();
            }

            myPage = new MyPage(this.Application, doc.Name, pageExplorer);
            myRibbonTracer.SetMyPage(myPage);
            this.Application.ShapeChanged += new Visio.EApplication_ShapeChangedEventHandler(myPage.onShapeChanged);
            this.Application.ShapeAdded += new Visio.EApplication_ShapeAddedEventHandler(myPage.onShapeAdded);
            myPage.start();

            pageExplorer.UpdateExplorer(doc);

        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            myRibbonTracer = new MyRibbonTracer();
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
