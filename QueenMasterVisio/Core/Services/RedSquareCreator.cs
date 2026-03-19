using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.Core.Services
{
    internal class RedSquareCreator
    {
        public static void RedSquareCreate(Page page)
        {
            // Если уже есть такой
            if (RedSquareGetLayer(page) != null)
                return;

            // Слой с красным квадратиком
            Visio.Layer redLayer = page.GetOrCreateLayer("RedLayer");
            // Основной слой
            Visio.Layer mainLayer = page.GetOrCreateLayer("Main");
            // Добавляем в основной слой
            foreach (Visio.Shape shape in page.Shapes)
            {
                mainLayer.Add(shape, 1);
                shape.Release();
            }

            double pageWidth = page.PageSheet.CellsU["PageWidth"].ResultIU;
            double pageHeight = page.PageSheet.CellsU["PageHeight"].ResultIU;

            // Дюйм
            double offset = 2.0 * 0.0393701;

            double left = -offset;
            double right = pageWidth + offset;
            double bottom = -offset;
            double top = pageHeight + offset;

            // Создаем
            Visio.Shape borderShape = page.DrawRectangle(left, bottom, right, top);
            Visio.Shape borderText = page.DrawRectangle(left, top + 0.2, 4, top + 0.25);

            // Добавляем к слою
            redLayer.Add(borderShape, 0);
            redLayer.Add(borderText, 0);
            // Заливка
            borderShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillPattern].FormulaU = "0";
            // Цвет линии
            borderShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineColor].FormulaU = "RGB(255,51,0)";
            // Толщина линии
            borderShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLineWeight].FormulaU = "0.04 in";


            borderText.Text = "Страница заблокирована: " + DateTime.Today.ToString("d");
            // Заливка
            borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillPattern].FormulaU = "0";
            // Линия
            borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLinePattern].FormulaU = "0";
            // Цвет шрифта
            borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowCharacter, (short)Visio.VisCellIndices.visCharacterColor].FormulaU = "RGB(255,51,0)";
            // Размер шрифта
            borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowCharacter, (short)Visio.VisCellIndices.visCharacterSize].FormulaU = "16 pt";
            // По левому 
            borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionParagraph, (short)Visio.VisRowIndices.visRowParagraph, (short)Visio.VisCellIndices.visHorzAlign].FormulaU = "0";


            // Ну и очищаем, пиздец
            borderShape.Release();
            redLayer.Release();
            mainLayer.Release();
            borderText.Release();
        }
        public static void RedSquareDelete(Page page)
        {
            Layer layer = RedSquareGetLayer(page);
            layer.SetOptions(0,0,0,0,0,0);
            layer.Delete(1);
            layer.Release();
        }
        public static Layer RedSquareGetLayer(Page page)
        {
            foreach (Layer layer in page.Layers)
                if (layer.NameU.StartsWith("RedLayer"))
                    return layer;
            return null;
        }
    }
}
