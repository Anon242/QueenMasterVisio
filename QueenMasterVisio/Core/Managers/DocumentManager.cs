using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System;
using System.Text.RegularExpressions;
using Page = Microsoft.Office.Interop.Visio.Page;


namespace QueenMasterVisio.Core.Managers
{
    internal static class DocumentManager
    {
        private static readonly Random _random = new Random();

        /// <summary>
        /// Считаем количество планов в документе
        /// </summary>
        public static int GetCountPlansInDocument(Page page)
        {
            int i = 0;
            foreach (Page item in page.Document.Pages)
                if (item.HasCell("User.pageCode") && item.IsPlanPage())
                    i++;

            return i;
        }

        public static Page CreateNewPage(string name)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc == null)
                return null;
            Page newPage = doc.Pages.Add();

            // Если такой есть, то просто даем рандом число в имя
            if (GetPageByName(name) != null)
                    name += $".{_random.Next(1000, 9999):0000}";

            newPage.Name = name;
             
            return newPage;
        }

        /// <summary>
        /// Находит страницу в активном документе по точному имени (без учёта регистра)
        /// Возвращает null, если страницы с таким именем нет
        /// </summary>
        public static Page GetPageByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            if (doc == null)
                return null;

            foreach (Page p in doc.Pages)
            {
                if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))
                    return p;
            }

            return null;
        }

    }
}
