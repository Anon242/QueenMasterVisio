using System;
using Microsoft.Office.Interop.Visio;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace QueenMasterVisio
{
    internal class Tools
    {

        // Получаем значениеU из ячейки
        public static string CellValueGet(Visio.Shape shape, string cell)
        {
            try
            {
                return shape.CellsU[cell].ResultStr[Visio.VisUnitCodes.visNoCast];
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaGet " + ex);
                return null;
            }

        }

        public static string CellValueGet(Visio.Page page, string cell)
        {
            try
            {
                return page.PageSheet.CellsU[cell].ResultStr[Visio.VisUnitCodes.visNoCast];
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ERROR CellFormulaGet " + ex);
                return null;
            }

        }

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

        public static string[] ExtractGValues(string input)
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
