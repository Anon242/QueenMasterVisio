using Microsoft.Office.Interop.Visio;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Visio.Application;
using Page = Microsoft.Office.Interop.Visio.Page;

namespace QueenMasterVisio
{
    // Класс который собирает инфу со всех страниц, дает инфу о страницах, проверяет ошибки 
    // 1 Поток смотрит за всеми страницами (не нагруженно) 
    // 2 Поток смотрит за конретной страницой на которой мы
    // Этот класс предоставляет страницу на которой мы находимся, дает слушатель с страницей
    public class BackgroundWorker
    {
        private Thread allPagesWorker;
        private Thread thisPagesWorker;
        protected string docName;
        private static bool ispause = false;

        // Слушатель, передает страницу
        public event EventHandler<Page> OnChangedPage;

        Application app;
        Document doc;

        // Таймера
        private const int allPagesTimer = 3000; // Сколько даем задержки на 1 страницу проходя по всему документу 
        private const int allPagesObjectTimer = 80; // Сколько даем задержку между получениями объектов с страницы

        private string pageName = "";
        // Должны передать имя документа где будем смотреть
        // Должны передать app
        public BackgroundWorker(Application _app, Document _doc) 
        {
            app = _app;
            doc = _doc;
            docName = doc.Name;
        }

        public void start()
        {
            allPagesWorker = new Thread(allPagesWorkerThreadFunc);
            thisPagesWorker = new Thread(thisPagesWorkerThreadFunc);

            allPagesWorker.IsBackground = true;
            thisPagesWorker.IsBackground = true;

            // Перед стартом надо накопить инфу
            //allPagesWorker.Start();
            thisPagesWorker.Start();
        }
        // Стопим потоки
        public void stop()
        {
            if (allPagesWorker.IsAlive)
                allPagesWorker.Abort();
            if (thisPagesWorker.IsAlive)
                thisPagesWorker.Abort();
        }

        public static void pause()
        {
            ispause = true;
        }
        public static void resume()
        {
            ispause = false;
        }

        /* 
         * Поток который собирает инфу со всех страниц документа
         *      С каждой страницы мы должны собирать:
         *          Платы, айди (прошивочный), тип (имя)
         *          Клемы, названия 
         *          
         *          Ищем по Terminal и Board
         */


        private void allPagesWorkerThreadFunc()
        {
            while (true)
            {
                try
                {
                    if (doc.Name == docName)
                    {
                        foreach (Page page in doc.Pages)
                        {
                            // Если активная страница - пропускаем
                            if (page?.Name == pageName) continue;

                            // ЧЕК ПО УСТРОЙСТВАМ
                            Regex regex = new Regex(@"^G\d");
                            if (!regex.IsMatch(page.Name)) continue;
                            //string gCode = page.Name.Split(' ').First();
                            Debug.WriteLine(string.Join(",", extractGValues(page.Name)));
                            /*
                            foreach (Shape shape in page.Shapes)
                            {
                                if(shape.Name.IndexOf("Terminal") != -1)
                                {

                                }
                                else if (shape.Name.IndexOf("Board") != -1)
                                {

                                }

                                // Освобождаем shape 
                                Marshal.ReleaseComObject(shape);
                                Thread.Sleep(allPagesObjectTimer);
                            }
                            */
                            // Освобождаем page 
                            Marshal.ReleaseComObject(page);
                            Thread.Sleep(allPagesTimer);

                        }
                    }

                }
                catch
                {
                    
                }
            }
        }

        // Поток который собирает инфу с конкретной страницы
        [System.Diagnostics.DebuggerStepThrough]
        private void thisPagesWorkerThreadFunc()
        {
            while (true)
            {
                try
                {
                    if (app?.ActiveDocument.Name == docName && app?.ActivePage.Name != pageName && !ispause)
                    {
                        onPageChanged(app.ActivePage);
                        pageName = app.ActivePage.Name;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("COM ERROR: " + ex.Message);
                    pageName = "";
                    Thread.Sleep(1500);
                }
                Thread.Sleep(200);
            }

        }
        private void onPageChanged(Page page)
        {
            OnChangedPage?.Invoke(this, page);
        }

        private static string[] extractGValues(string input)
        {
            Regex regex = new Regex(@"G\d+(?:\.\d+)?\b");

            MatchCollection matches = regex.Matches(input);

            List<string> result = new List<string>();
            foreach (Match match in matches)
            {
                result.Add(match.Value);
            }

            return result.ToArray();
        }

    }
}
