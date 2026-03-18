using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.Core.Helpers
{
    /// <summary>
    /// Методы расширения для объектов Visio (Page, Shape, Layer)
    /// </summary>
    public static class VisioExtensions
    {
        #region === Page Extensions ===

        /// <summary>Проверяет существование User-ячейки</summary>
        public static bool HasCell(this Page page, string cellName)
        {
            return page.PageSheet.CellExistsU[cellName, (short)VisExistsFlags.visExistsAnywhere] != 0; 
        }

        /// <summary>Получить значение User-ячейки (строка)</summary>
        public static string GetCellFormulaU(this Page page, string cellName)
        {
            return page.HasCell(cellName)
                ? page.PageSheet.CellsU[cellName].FormulaU.Replace("\"", "")
                : string.Empty;
        }

        /*
        public static string GetCellResultU(this Page page, string cellName)
        {
            return page.HasCell(cellName)
                ? page.PageSheet.CellsU[cellName].FormulaU.Replace("\"", "")
                : string.Empty;
        }
        */

        private static string _GetPlanCode(this Page page)
        {
            if (!page.IsPlanPage())
                return string.Empty;

            foreach (Layer layer in page.Layers)
            {
                string celIndex = layer.Index == 0 ? "" : '[' + "" + (layer.Index) + ']';
                if (page.PageSheet.Cells[$"Layers.Active" + celIndex].Formula == "1")
                {
                    return layer.Name;
                }

            }
            return "All";
        }

        /// <summary>Установить User-ячейку (создаёт автоматически, если нет)</summary>
        public static void SetUserCell(this Page page, string cellName, string value)
        {
            const short section = (short)VisSectionIndices.visSectionUser;
            if (!page.HasCell(cellName))
            {
                short row = (short)page.PageSheet.AddNamedRow(section, cellName, (short)VisRowTags.visTagDefault);
            }
            page.PageSheet.CellsU[$"User.{cellName}"].FormulaU = $"\"{value}\"";
        }

        /// <summary>Быстрые проверки типа страницы</summary>
        public static bool IsPlanPage(this Page page) => page.GetCellFormulaU("pageCode") == "Plan";
        public static bool IsAutoTracePage(this Page page) => page.GetCellFormulaU("pageCode") == "planAuto";
        public static bool IsBackgroundPage(this Page page) => page.Background != 0;

        /// <summary>Получить код страницы (аналог твоего getPageCode)</summary>
        public static string GetPageCode(this Page page) => page.GetCellFormulaU("pageCode");
        public static string GetPlanCode(this Page page) => page._GetPlanCode();

        /// <summary>Заблокировать/разблокировать все слои</summary>
        public static void LockAllLayers(this Page page, bool lockState = true)
        {
            for (short i = 0; i < page.Layers.Count; i++)
            {
                string index = i == 0 ? "" : "[" + (i + 1) + "]";
                page.PageSheet.Cells[$"Layers.Locked{index}"].Formula = lockState ? "1" : "0";
                page.PageSheet.Cells[$"Layers.Active{index}"].Formula = "0";
            }
        }

        /// <summary>Включить/выключить печать всех слоёв (кроме Plan)</summary>
        public static void SetPrintOnLayers(this Page page, bool print)
        {
            for (short i = 0; i < page.Layers.Count; i++)
            {
                string index = i == 0 ? "" : "[" + (i + 1) + "]";
                Layer layer = page.Layers[i + 1];
                if (layer.Name != "Plan")
                    page.PageSheet.Cells[$"Layers.Print{index}"].Formula = print ? "1" : "0";
            }
        }

        #endregion

        #region === Shape Extensions ===

        /// <summary>Это соединительная линия?</summary>
        public static bool IsLine(this Shape shape)
        {
            return shape.Master?.NameU.Contains("Динамическая соединительная линия") == true ||
                   shape.Name.Contains("Dynamic connector") ||
                   Checker.isLine(shape); // можно оставить твой Checker пока
        }

        /// <summary>Это устройство/щит/свет и т.д.</summary>
        public static bool IsDevice(this Shape shape)
            => shape.Name.Contains("Device") || shape.Name.Contains("Shield") ||
               shape.Name.Contains("Light") || shape.Name.Contains("Camera");

        /// <summary>Установить формулу в любую ячейку (удобнее CellFormulaSet)</summary>
        public static void SetFormula(this Shape shape, string cellName, string formula)
        {
            shape.CellsU[cellName].FormulaU = formula;
        }

        /// <summary>Получить значение свойства (Prop.)</summary>
        public static string GetProp(this Shape shape, string propName)
        {
            if (shape.CellExistsU[$"Prop.{propName}", (short)VisExistsFlags.visExistsAnywhere] == 0)
                return string.Empty;
            return shape.CellsU[$"Prop.{propName}"].FormulaU.Replace("\"", "");
        }

        /// <summary>Приклеить начало линии к фигуре</summary>
        public static void GlueBeginTo(this Shape connector, Shape target)
        {
            connector.CellsU["BeginX"].GlueTo(target.CellsSRC[(short)VisSectionIndices.visSectionConnectionPts, 0, (short)VisCellIndices.visX]);
            connector.CellsU["BeginY"].GlueTo(target.CellsSRC[(short)VisSectionIndices.visSectionConnectionPts, 0, (short)VisCellIndices.visY]);
        }

        /// <summary>Приклеить конец линии к фигуре</summary>
        public static void GlueEndTo(this Shape connector, Shape target)
        {
            connector.CellsU["EndX"].GlueTo(target.CellsSRC[(short)VisSectionIndices.visSectionConnectionPts, 0, (short)VisCellIndices.visX]);
            connector.CellsU["EndY"].GlueTo(target.CellsSRC[(short)VisSectionIndices.visSectionConnectionPts, 0, (short)VisCellIndices.visY]);
        }

        /// <summary>Безопасно удалить фигуру (снимает блокировки)</summary>
        public static void SafeDelete(this Shape shape)
        {
            try
            {
                shape.CellsU["LockDelete"].FormulaU = "0";
                shape.Delete();
            }
            catch { }
        }

        #endregion

        #region === Layer Extensions ===

        /// <summary>Установить видимость, печать, блокировку и т.д. одной строкой</summary>
        public static void SetOptions(this Layer layer, bool visible, bool print, bool active, bool locked, bool snap, bool glue)
        {
            Page page = layer.Page;
            string index = layer.Index == 0 ? "" : "[" + layer.Index + "]";

            page.PageSheet.Cells[$"Layers.Visible{index}"].Formula = visible ? "1" : "0";
            page.PageSheet.Cells[$"Layers.Print{index}"].Formula = print ? "1" : "0";
            page.PageSheet.Cells[$"Layers.Active{index}"].Formula = active ? "1" : "0";
            page.PageSheet.Cells[$"Layers.Locked{index}"].Formula = locked ? "1" : "0";
            page.PageSheet.Cells[$"Layers.Snap{index}"].Formula = snap ? "1" : "0";
            page.PageSheet.Cells[$"Layers.Glue{index}"].Formula = glue ? "1" : "0";
        }

        /// <summary>Добавить фигуру на слой безопасно</summary>
        public static void AddShape(this Layer layer, Shape shape, bool allow1D = true)
        {
            try { layer.Add(shape, allow1D ? (short)1 : (short)0); }
            catch { }
        }

        #endregion

        #region === Общие COM-хелперы ===

        /// <summary>Безопасно освободить COM-объект (можно вызывать всегда)</summary>
        public static void Release(this object comObject)
        {
            if (comObject != null)
            {
                try { Marshal.ReleaseComObject(comObject); }
                catch { }
            }
        }

        #endregion
    }
}
