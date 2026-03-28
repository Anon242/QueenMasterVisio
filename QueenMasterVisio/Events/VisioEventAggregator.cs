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

        public VisioEventAggregator(Application _app, Explorer _explorer)
		{
			app = _app;
            explorer = _explorer;
        }


        public void start()
		{
			// Старт BackgroundWirker'a
			backgroundWorker = new BackgroundWorker(app, app.ActiveDocument);
			backgroundWorker.OnChangedPage += OnPageChanged;
			backgroundWorker.start();
		}

		private void OnPageChanged(object sender, Page page)
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
            if (explorer == null)
                return;
            
            explorer.LookDevices(currentPage);
        }

        public void OnShapeChanged(Visio.Shape shape)
		{
			//Debug.WriteLine("Shape changed: " + shape.Name);
		}

		public void OnShapeAdded(Visio.Shape shape)
		{
            // Не вызываем если хоть кто то запретил это делать или есть флаг undo (ctrl+z | ctrl+y)
            if(VisioEventSuppressor.IsShapeAddedSuppressed || shape.Application.IsUndoingOrRedoing)
				return;

            if (shape.IsLine())
			{
                int scopeId = Globals.ThisAddIn.Application.BeginUndoScope("Изменение линии");  /////<

                if (shape.ContainingPage.IsPlanPage()) 
				{
                    ShapeManager.RebuildShapePlan(shape);
				}
                // Предположим что тут у нас девайсы
                else
                {
                    ShapeManager.RebuildShapeDevice(shape);
                }

                Globals.ThisAddIn.Application.EndUndoScope(scopeId, true);                      /////>
            }
        }
	}
}
