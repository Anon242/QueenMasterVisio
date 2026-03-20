using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Managers;
using QueenMasterVisio.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio
{

	public class VisioEventAggregator
	{
		public static Application app;
		public Collection<string> whiteList = new Collection<string>() { "QueenCallout" };
		public static Explorer explorer;
        BackgroundWorker backgroundWorker;

        public VisioEventAggregator(Application _app, string _name, Explorer _explorer)
		{
			app = _app;
            explorer = _explorer;
        }


        public void start()
		{
			// Старт BackgroundWirker'a
			backgroundWorker = new BackgroundWorker(app, app.ActiveDocument);
			backgroundWorker.OnChangedPage += onPageChanged;
			backgroundWorker.start();
		}

		private void onPageChanged(object sender, Page page)
		{
			//page.Application.DoCmd((short)VisUICmds.visCmdViewFitInWindow); // Команда выровнять по ширине

			if (page.IsPlanPage())
			{
                Debug.WriteLine("Мы на плане " + page.NameU);
                MainLentXml.UpdateLayerButtons(page.GetPlanCode());
                MainLentXml.RibbonReload(true);
			}
			else
			{
				MainLentXml.RibbonReload(false);
                // Тут мы применяем настройки
                // LineAdjustFrom 1
                // LineAdjustTo 2
                Tools.CellFormulaSet(page, "LineAdjustFrom", "1");
                Tools.CellFormulaSet(page, "LineAdjustTo", "2");
                Tools.CellFormulaSet(page, "RouteStyle", "17");
            }

        }

        public void SomeMethod()
        {
            Visio.Page currentPage = app.ActivePage;

            // Вызываем метод LookDevices у экземпляра explorer
            if (explorer != null)
            {
                explorer.LookDevices(currentPage);
            }
        }

		public void onShapeChanged(Visio.Shape shape)
		{
			//Debug.WriteLine("Shape changed: " + shape.Name);
		}

		public void onShapeAdded(Visio.Shape shape)
		{
            // Не вызываем если хоть кто то запретил это делать или есть флаг undo
            if(VisioEventSuppressor.IsShapeAddedSuppressed || shape.Application.IsUndoingOrRedoing)
				return;
            if (shape.IsLine())
			{
                int scopeId = Globals.ThisAddIn.Application.BeginUndoScope("Изменение линии");

                if (shape.ContainingPage.IsPlanPage()) 
				{
					rebuildShape(shape, true);
				}
                // Предположим что тут у нас девайсы
                else
                {
					rebuildShapeDevice(shape);
                }
                Globals.ThisAddIn.Application.EndUndoScope(scopeId, true);

            }

        }

        private static void rebuildShape(Visio.Shape shape, bool onUserShape = false)
        {
            string activePlanCode = shape.ContainingPage.GetPlanCode();  
            if (shape.ContainingPage.GetPlanCode() == null)
            {
                shape.Delete();
                return;
            }
            // Если она ничкему не присоеденена
            if (shape.Connects.Count != 2)
            {
                shape.Delete();
                return;
            }
            // На плане не даем размещать линии
            if (shape.ContainingPage.GetPlanCode() == "Plan")
            {
                shape.Delete();
                return;
            }
            // Если мы на трасировке,добавляем юзершейпы на линии
            else if (WireService.wires.Any(w => w.name.Contains(activePlanCode)&& w.isWire)) 
            {

                Visio.Shape connectedShapeFrom = shape.Connects[1].ToSheet;
                Visio.Shape connectedShapeTo = shape.Connects[2].ToSheet;
            }

            // Логика от и до (test)
            SetIdInLine(shape);

            shape.CellsU["Rounding"].FormulaU = "2 mm";
            try
            {
                shape.CellsU["LineWeight"].FormulaU = "IFERROR(1.5 pt * ThePage!Prop.Scale,1 pt)";
            }
            catch
            {
                shape.CellsU["LineWeight"].FormulaU = "1.5 pt";

            }
            shape.CellsU["LineColor"].FormulaU = WireService.GetWireByName(activePlanCode).color;

            shape.ContainingPage.Layers[activePlanCode].Add(shape, 0);
            shape.SendToBack();
            shape.ContainingPage.Layers["Соединительная линия"].CellsC[(short)VisCellIndices.visLayerPrint].FormulaU = "0";

            shape.CellsU["LockBegin"].FormulaU = "1";
            shape.CellsU["LockEnd"].FormulaU = "1";
            shape.CellsU["LockMoveX"].FormulaU = "1";
            shape.CellsU["LockMoveY"].FormulaU = "1";
            shape.CellsU["LockRotate"].FormulaU = "1";
            shape.CellsU["ShapeSplittable"].FormulaU = "0";
            shape.CellsU["ConFixedCode"].FormulaU = "2";
            shape.CellsU["ConLineRouteExt"].FormulaU = "0";
            shape.CellsU["ConLineRouteExt"].FormulaU = "1";


            Visio.Shape nearestLine = FindNearestLine(shape);
            if (nearestLine != null)
                MergeLineGeometry(nearestLine, shape);




        }
        public static void rebuildShapeDevice(Shape shape)
		{
            // Линия уже перекрашена, не трогаем
            if (shape.GetCellFormulaU("ConFixedCode") == "2")
                return;

            if (shape.ContainingPage.HasCell("Prop.Scale"))
			{
				shape.CellsU["LineWeight"].FormulaU = "=IFERROR(ThePage!Prop.Scale*0.5&\"pt\",1)";
            }

            foreach (Visio.Shape candidate in shape.ContainingPage.Shapes.Cast<Shape>().Reverse())
            {
                if (candidate.IsLine() && candidate.ID != shape.ID)
                {

                    if ((candidate.CellsU["BeginX"].FormulaU == shape.CellsU["BeginX"].FormulaU && candidate.CellsU["BeginY"].FormulaU == shape.CellsU["BeginY"].FormulaU) ||
                        (candidate.CellsU["EndX"].FormulaU == shape.CellsU["EndX"].FormulaU && candidate.CellsU["EndY"].FormulaU == shape.CellsU["EndY"].FormulaU) ||
                        (candidate.CellsU["BeginX"].FormulaU == shape.CellsU["EndX"].FormulaU && candidate.CellsU["BeginY"].FormulaU == shape.CellsU["EndY"].FormulaU) ||
                        (candidate.CellsU["EndX"].FormulaU == shape.CellsU["BeginX"].FormulaU && candidate.CellsU["EndY"].FormulaU == shape.CellsU["BeginY"].FormulaU))
                    {

                        if (candidate.CellsU["LineColor"].FormulaU != "0" && candidate.CellsU["LineColor"].FormulaU != "\"RGB(0;0;0)\"")
                        {
                            shape.CellsU["LineColor"].FormulaU = candidate.CellsU["LineColor"].FormulaU;
                            shape.CellsU["LinePattern"].FormulaU = candidate.CellsU["LinePattern"].FormulaU;
                            shape.CellsU["LineWeight"].FormulaU = candidate.CellsU["LineWeight"].FormulaU;
                            break;
                        }

                    }
                }
            }

            shape.CellsU["Rounding"].FormulaU = "3 mm";
            shape.CellsU["ShapeRouteStyle"].FormulaU = "17";
            shape.CellsU["ConFixedCode"].FormulaU = "2";

            //shape.Application.QueueMarkerEvent("Recalc"); // заставляет Visio пересчитать
            //shape.CellsU["ObjType"].FormulaU = "4";


            // shape.CellsU["Path"].FormulaU = "";

        }


        private static Visio.Shape FindNearestLine(Visio.Shape line)
		{
			// Получаем его конец и ищем ближайшую, вернем shape близкого
			// Радиус поиска нужен маленькой
			double endX = line.CellsU["EndX"].ResultIU;
			double endY = line.CellsU["EndY"].ResultIU;
            double beignX = line.CellsU["BeginX"].ResultIU;
            double beignY = line.CellsU["BeginY"].ResultIU;

            //Wire wire = MyPage.wires.FirstOrDefault(w => w.color == line.CellsU["LineColor"].FormulaU);

			double min = 999;
            double oldMin = min;

            Visio.Shape bufferShapeline = null;

            foreach (Visio.Shape shape in line.Application.ActivePage.Shapes.Cast<Visio.Shape>().Reverse())
			{
                if (shape.IsLine())
                {
                    //Wire wireShape = MyPage.wires.FirstOrDefault(w => w.color == shape.CellsU["LineColor"].FormulaU);

                    double endXShape = shape.CellsU["EndX"].ResultIU;
                    double endYShape = shape.CellsU["EndY"].ResultIU;
                    double beignXShape = shape.CellsU["BeginX"].ResultIU;
                    double beignYShape = shape.CellsU["BeginY"].ResultIU;

                    // Начало у них одинаковое и не она же сама
                    if (beignX == beignXShape && beignY == beignYShape && shape.NameU != line.NameU)
                    {
                        min = Math.Min(Tools.LineLenght(endX, endY, endXShape, endYShape), min);
                        if (min != oldMin)
                        {
                            bufferShapeline = shape;
                        }
                        oldMin = min;
                    }
                }
            }
			return bufferShapeline;
        }
            

        private static void MergeLineGeometry(Visio.Shape fromLine, Visio.Shape toLine)
		{
            //Visio.Shape nearestLine = FindNearestLine()
            // Запомним последний элемент у to, все стираем
            // Копируем все кроме последнего у from и вставим в to с последним что мы запомнили
            short geometrySection = (short)Visio.VisSectionIndices.visSectionFirstComponent;

            Visio.Section fromGeometry = fromLine.Section[(short)Visio.VisSectionIndices.visSectionFirstComponent];
            Visio.Section toGeometry = toLine.Section[(short)Visio.VisSectionIndices.visSectionFirstComponent];

            // Если линия тупо прямая
            //if (toGeometry.Count <= 2)
            //	return;
            // Запомнили последнее
            Visio.Row lastRowtoGeometry = toGeometry[(short)(toGeometry.Count -1)];
			
            string x = toLine.CellsSRC[geometrySection, lastRowtoGeometry.Index, (short)Visio.VisCellIndices.visX].FormulaU;
            string y = toLine.CellsSRC[geometrySection, lastRowtoGeometry.Index, (short)Visio.VisCellIndices.visY].FormulaU;

            //Удаляем все элементы кроме последнего
            while (toGeometry.Count > 1)
            {
                toLine.DeleteRow(geometrySection, 1); // Удаляем с начала
            }

			string test = fromLine.CellsSRC[geometrySection, 2, (short)Visio.VisCellIndices.visX].FormulaU;

			
            for (short i = 1; i <= fromGeometry.Count-1; i++)
            {
                // Получаем тип строки из исходной геометрии
                short rowType = fromLine.RowType[geometrySection, i];
				
                // Добавляем строку того же типа в целевую геометрию
                short newRowIndex = toLine.AddRow(geometrySection, i, (short)rowType);

                // Копируем значения ячеек
                toLine.CellsSRC[geometrySection, newRowIndex, (short)Visio.VisCellIndices.visX].FormulaU =
                    fromLine.CellsSRC[geometrySection, i, (short)Visio.VisCellIndices.visX].FormulaU;

                toLine.CellsSRC[geometrySection, newRowIndex, (short)Visio.VisCellIndices.visY].FormulaU =
                    fromLine.CellsSRC[geometrySection, i, (short)Visio.VisCellIndices.visY].FormulaU;
            }

            // Добавляем строку того же типа в целевую геометрию
            short newRowIndex1 = toLine.AddRow(geometrySection, (short)(toGeometry.Count), (short)139);
			// Копируем значения ячеек
			toLine.CellsSRC[geometrySection, newRowIndex1, (short)Visio.VisCellIndices.visX].FormulaU = x;
			toLine.CellsSRC[geometrySection, newRowIndex1, (short)Visio.VisCellIndices.visY].FormulaU = y;



            if (toGeometry.Count >= 3)
			{

                string x1 = toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 3)].Index, (short)Visio.VisCellIndices.visX].FormulaU;
                string x2 = toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 2)].Index, (short)Visio.VisCellIndices.visX].FormulaU;

				// y короче чем x
				if (x1 == x2)
				{
					toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 2)].Index, (short)Visio.VisCellIndices.visY].FormulaU =
						toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 1)].Index, (short)Visio.VisCellIndices.visY].FormulaU;
                }
				else
				{
					toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 2)].Index, (short)Visio.VisCellIndices.visX].FormulaU =
						toLine.CellsSRC[geometrySection, toGeometry[(short)(toGeometry.Count - 1)].Index, (short)Visio.VisCellIndices.visX].FormulaU;
                }
            }



            // Восстанавливаем последний элемент
            //toGeometry[1].CellsU["X"].FormulaU = lastX.ToString();
            //toGeometry[1].CellsU["Y"].FormulaU = lastY.ToString();
            //toGeometry[1].CellsU["Type"].FormulaU = lastType.ToString();

        }



        /*
		   foreach (Connect shape in lineShape.Connects)
			{
				Debug.WriteLine(shape.ToSheet.Name);
	
			}
		 */

        // Тут мы определяем и назначаем айди линии на трасировочных слоях плана
        // Нужен рекурсивный метод
        private static void SetIdInLine(Visio.Shape lineShape)
		{
			Visio.Shape connectedShape = lineShape.Connects[2].ToSheet;
            string activePlanCode = lineShape.ContainingPage.GetPlanCode();

            if (connectedShape.Name.Contains("Device") || connectedShape.Name.Contains("Light") || connectedShape.Name.Contains("Camera") || connectedShape.Name.Contains("Alarm") || connectedShape.Name.Contains("Sound") || connectedShape.Name.Contains("Box"))
            {
				if (connectedShape.CellExists["Prop.Number", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0)
				{
                    short id = (short)lineShape.ID;
                    string nameValue = connectedShape.CellsU["Prop.Number"].FormulaU.Replace("\"", "");

                    //try
                    //{
                    Visio.Master master = lineShape.Document.Masters["QueenCallout"];
                    Visio.Shape shape = lineShape.Document.Application.ActivePage.Drop(master, (double)lineShape.CellsU["EndX"].ResultIU, (double)lineShape.CellsU["EndY"].ResultIU);

                    shape.Text = activePlanCode[0] + nameValue;
					// Если распределительная коробка 
					if(connectedShape.Name.Contains("Box"))
                        shape.Text = activePlanCode[0] +"J" + nameValue;

                    
                    string nameLineId = "Sheet." + lineShape.ID;

                    /*
                        // Назначаем формулы
                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "BeginX", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"{nameLineId}!BeginX";
                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "BeginY", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"{nameLineId}!BeginY";

                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "EndX", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"{nameLineId}!EndX";
                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "EndY", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"{nameLineId}!EndY";

                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "LineDX", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"User.EndX - User.BeginX";
                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "LineDY", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"User.EndY - User.BeginY";

                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "LineLengthSq", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"User.LineDX * User.LineDX + User.LineDY * User.LineDY";

                        row = shape.AddNamedRow((short)VisSectionIndices.visSectionUser, "T", (short)VisRowTags.visTagDefault);
                        shape.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"IF(User.LineLengthSq=0, 0, MIN(1, MAX(0, ((Controls.End.X - User.BeginX) * User.LineDX + (Controls.End.Y - User.BeginY) * User.LineDY) / User.LineLengthSq)))";

                        shape.CellsU["PinX"].Formula = $"GUARD(User.BeginX + User.T * User.LineDX)";
                        shape.CellsU["PinY"].Formula = $"GUARD(User.BeginY + User.T * User.LineDY)";
					*/


                    shape.CellsU["Char.Color"].FormulaU = WireService.GetWireByName(activePlanCode).color;
     //               }
					//catch
					//{
					//	Debug.WriteLine("В образах документа остуствует QueenCallout");
					//}
				}
			}
		}

        private static List<(double x, double y)> GetGeometryCoordinates(Visio.Shape shape)
        {
            var coordinates = new List<(double x, double y)>();

            try
            {
                Visio.Section geometrySection = shape.Section[(short)Visio.VisSectionIndices.visSectionFirstComponent];

                for (short row = 0; row < geometrySection.Shape.RowCount[(short)Visio.VisSectionIndices.visSectionFirstComponent]; row++)
                {
                    try
                    {
                        double x = shape.CellsSRC[
                            (short)Visio.VisSectionIndices.visSectionFirstComponent,
                            row,
                            (short)Visio.VisCellIndices.visX].ResultIU;

                        double y = shape.CellsSRC[
                            (short)Visio.VisSectionIndices.visSectionFirstComponent,
                            row,
                            (short)Visio.VisCellIndices.visY].ResultIU;

                        coordinates.Add((x, y));
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка: {ex.Message}");
            }

            return coordinates;
        }
	}
}
