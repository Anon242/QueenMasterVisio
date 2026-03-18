using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;
namespace QueenMasterVisio.Core.Services
{
    static class WireAutoConnectionService
    {
        public static void autoConnect(Page page)
        {
            if (!page.IsPlanPage())
                return;

            string activePlanCode = page.GetPlanCode();

            Visio.Shape shield = null;
            List<Visio.Shape> devices = new List<Visio.Shape>();

            if (!WireService.wires.Any(w => w.name == activePlanCode && w.isWire))
            {
                MessageBox.Show("Ошибка слоя", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selection = page.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper,
            page.Layers.ItemU[activePlanCode]);

            // Ищем фигуры
            foreach (Visio.Shape shape in selection)
            {
                if (shape.Name.Contains("Device") || shape.Name.Contains("Light") || shape.Name.Contains("Camera") || shape.Name.Contains("Sound"))
                {
                    devices.Add(shape);
                }
            }
            selection.DeselectAll();
            // Теперь ищем щит
            foreach (Visio.Shape shape in page.Shapes)
                if (shape.Name.Contains("Shield"))
                    shield = shape;


            if (shield == null)
            {
                MessageBox.Show("Добавьте 1 щит из фигур",
               "Щит не найден",
               MessageBoxButtons.OK,
               MessageBoxIcon.Warning);
                return;
            }
            else if (devices.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы 1 девайс на план",
              "Не найден ни один девайс на плане",
              MessageBoxButtons.OK,
              MessageBoxIcon.Warning);
                return;
            }

            foreach (Visio.Shape shape in devices)
            {
                if (!checkConnetctedLinesInDevice(page, shape))
                {
                    CreateAndGlueConnector(page, shield, shape, activePlanCode);
                }
            }
        }

        private static bool checkConnetctedLinesInDevice(Page page, Visio.Shape shape)
        {
            foreach (Visio.Shape connector in page.Shapes.Cast<Visio.Shape>().Where(s => CheckerService.isLine(s)))
            {
                if (connector.Layer[2].Name == page.GetPlanCode())
                {
                    Visio.Shape beginShape = connector.Connects[1].ToSheet;  // Индекс 1 = Begin
                    Visio.Shape endShape = connector.Connects[2].ToSheet;    // Индекс 2 = End

                    // Проверяем, подключен ли коннектор к целевой фигуре
                    if (beginShape == shape || endShape == shape)
                        return true;
                }
            }
            return false;
        }

        public static bool rebuildBrake = false;
        public static void CreateAndGlueConnector(Page page, Visio.Shape fromShape, Visio.Shape toShape, string layerName)
        {
            // Костыль
            rebuildBrake = true;
            Visio.Shape connector = CreateConnector(page);
            Debug.WriteLine(connector.Name);
            GlueConnectorToShapes(connector, fromShape, toShape);
            rebuildBrake = false;
            //rebuildShape(connector, true);

            //Marshal.ReleaseComObject(connector);
        }
        // Костыль


        private static Visio.Shape CreateConnector(Page page)
        {
            Master connectorMaster = page.Document.Masters["Динамическая соединительная линия"];
            Visio.Shape connector = page.Drop(connectorMaster, 5, 5);
            //Marshal.ReleaseComObject(connectorMaster);
            return connector;
        }

        private static void GlueConnectorToShapes(Visio.Shape connector, Visio.Shape fromShape, Visio.Shape toShape)
        {
            // Получаем первую точку соединения исходной фигуры
            Cell fromPointX = fromShape.CellsSRC[
                (short)VisSectionIndices.visSectionConnectionPts,
                (short)VisRowIndices.visRowConnectionPts,
                (short)VisCellIndices.visX];

            Cell fromPointY = fromShape.CellsSRC[
                (short)VisSectionIndices.visSectionConnectionPts,
                (short)VisRowIndices.visRowConnectionPts,
                (short)VisCellIndices.visY];

            // Получаем первую точку соединения целевой фигуры
            Cell toPointX = toShape.CellsSRC[
                (short)VisSectionIndices.visSectionConnectionPts,
                (short)VisRowIndices.visRowConnectionPts,
                (short)VisCellIndices.visX];

            Cell toPointY = toShape.CellsSRC[
                (short)VisSectionIndices.visSectionConnectionPts,
                (short)VisRowIndices.visRowConnectionPts,
                (short)VisCellIndices.visY];

            connector.CellsU["BeginX"].GlueTo(fromPointX);
            connector.CellsU["BeginY"].GlueTo(fromPointY);

            connector.CellsU["EndX"].GlueTo(toPointX);
            connector.CellsU["EndY"].GlueTo(toPointY);
        }



    }

}
