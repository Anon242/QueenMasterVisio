using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Visio.Page;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Reflection.Emit;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.ObjectModel;

namespace QueenMasterVisio
{
    internal class Tools
    {
        // Получаем формулуU из ячейки
        public static string CellFormulaGet(Visio.Shape shape, string cell)
        {
            try
            {
                return shape.CellsU[cell].FormulaU;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaGet " + ex);
                return null;
            }
            
        }
        // Перегружен Получаем формулуU из ячейки в странице
        public static string CellFormulaGet(Visio.Page page, string cell)
        {
            try
            {
                return page.PageSheet.CellsU[cell].FormulaU.Replace("\"", "");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaGet " + ex);
                return null;

            }

        }
        // Назначаем формулуU в ячейку
        public static void CellFormulaSet(Visio.Shape shape, string cell, string set)
        {
            try
            {
                shape.CellsU[cell].FormulaU = set;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaSet " + ex);
                throw;

            }
        }
        // Перегружен Назначаем формулуU в ячейку страницы
        public static void CellFormulaSet(Visio.Page page, string cell, string set)
        {
            try
            {
                page.PageSheet.CellsU[cell].FormulaU = set;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaSet " + ex);
                throw;

            }
        }
        // Создание User.* 
        public static void CellUserCreate(Visio.Shape shape, string cell, string formula)
        {
            try
            {
                short row = (short)shape.AddNamedRow((short)VisSectionIndices.visSectionUser, cell, (short)VisRowTags.visTagDefault);
                shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = formula;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellUserCreate " + ex);
                throw;

            }
        }
        // Существует ли ячейка у shape
        public static bool CellExistsCheck(Visio.Shape shape, string cell)
        {
            return shape.CellExists[cell, (short)Visio.VisExistsFlags.visExistsAnywhere] != 0;
        }
        // Существует ли ячейка у page.PageSheet
        public static bool CellExistsCheck(Visio.Page page, string cell)
        {
            return page.PageSheet.CellExists[cell, (short)Visio.VisExistsFlags.visExistsAnywhere] != 0;
        }

        public static double LineLenght(double px1, double py1, double px2, double py2)
        {
            return Math.Sqrt(Math.Pow(px1 - px2, 2) + Math.Pow(py1 - py2, 2));
        }
    }
}
