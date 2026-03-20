using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Services;
using QueenMasterVisio.Ribbon;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.Core.Managers
{
    internal class PageManager
    {

        /// <summary>
        /// Когда мы на любой другой странице, мы можем создать план 
        /// </summary>
        /// <param name="page"></param>
        public static void CreateNewPlan(Page page)
        {

            // Кол во планов в документе
            int planCount = DocumentManager.GetCountPlansInDocument(page); 
            int id = planCount;

            // Создаем новую страницу
            Page newPage = DocumentManager.CreateNewPage(VisioEventAggregator.explorer.ShowRenameDialog("Plan." +id));
           
            newPage.SetUserCell("pageCode", "Plan");
            newPage.SetUserCell("id", id.ToString());
            newPage.SetPropCell("Scale", 1.ToString(), "Размер элементов");

            // Создаем все слои
            foreach (var item in WireService.wires)
            {
                newPage.Layers.Add(item.name);
            }

            // Запрещаем печать все кроме первой страницы
            newPage.SetPrintOnLayersNoPlan(false);
            SetOptionsAllPlanlayer(newPage, "Plan");
            // Вызываем перерисовку Ленты
            MainLentXml.RibbonReload(true);

        }
        /// <summary>
        /// Настраиваем слои на плане
        /// </summary>
        /// <param name="page"></param>
        /// <param name="planCode"></param>
        public static void SetOptionsAllPlanlayer(Page page, string planCode)
        {
            if (!page.IsPlanPage())
                return;

            planCode = planCode.Replace("btn", "");

            foreach (Layer layer in page.Layers)
            {
                if (planCode == "All")
                {
                    layer.SetOptions(1, 1, 0, 1, 0, 0);
                    continue;
                }
                // Обнуляем слой, следующими условиями все настроим
                layer.SetOptions(0, 0, 0, 1, 0, 0);
                // Px, Cx, ...
                if (layer.Name.Contains(planCode))
                    layer.SetOptions(1, 0, 1, 0, 1, 1);
                // Мы не на плане, но план должен быть отображен
                if (layer.Name == "Plan" && planCode != "Plan")
                    layer.SetOptions(1, 1, 0, 1, 1, 1);
                // Если мы на плане
                else if (layer.Name == "Plan" && planCode == "Plan")
                    layer.SetOptions(1, 1, 1, 0, 1, 1);
                // Слой с линиями просто скрываем, он не нужен
                if (layer.Name == "Соединительная линия")
                    layer.SetOptions(0, 0, 0, 0, 0, 0);
                if (layer.Name == "Develop")
                    layer.SetOptions(1, 0, 0, 0, 0, 0);
                layer.Release();
            }
            MainLentXml.UpdateLayerButtons(planCode);
            MainLentXml.RibbonReload(true);

        }

        public static void LockOrUnlockLayer(Page page)
        {
            if (page.IsPlanPage() || page.IsAutoTracePage())
                return;

            // Проверим заблокана ли стр
            int lockCount = 0;
            foreach (Layer layer in page.Layers)
            {
                string index = layer.Index == 0 ? "" : "[" + layer.Index + "]";
                if (page.PageSheet.Cells[$"Layers.Locked" + index].FormulaU.Replace("\"", "") == "1")
                {
                    lockCount++;
                }
                // Страница заблокирована 
                if(lockCount == page.Layers.Count)
                {
                    RedSquareCreator.RedSquareDelete(page);
                    LockAllLayers(page, false);
                }
                // Страница разблокирована 
                else
                {
                    RedSquareCreator.RedSquareCreate(page);
                    LockAllLayers(page, true);
                }
            }
            
            
        }
        /// <summary>Заблокировать/разблокировать все слои</summary>
        public static void LockAllLayers(Page page, bool lockState = true)
        {
            for (short i = 0; i < page.Layers.Count; i++)
            {
                string index = i == 0 ? "" : "[" + (i + 1) + "]";
                page.PageSheet.Cells[$"Layers.Locked{index}"].Formula = lockState ? "1" : "0";
                page.PageSheet.Cells[$"Layers.Active{index}"].Formula = "0";
            }
        }

        /// <summary>
        /// Находим Page к которому он привязан, берем его слой и перерисовываем
        /// </summary>
        /// <param name="page"></param>
        public static void RedrawPageAuto(Page page)
        {
            /*
            string pageType = page.PageSheet.CellsU["User.pageType"].Formula;
            string id = page.PageSheet.CellsU["User.id"].Formula;

            Page sPage = searchPlan(page, id);
            clearAlShapesInPage(page);

            if (sPage != null)
                copyPasteInLayer(sPage, page, GetOrCreateLayer(sPage, pageType));
            */
        }
    }
}
