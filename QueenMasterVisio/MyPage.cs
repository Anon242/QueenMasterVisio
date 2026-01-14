using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using Application = Microsoft.Office.Interop.Visio.Application;
using Office = Microsoft.Office.Core;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio
{

	public class MyPage
	{
		Application app;
		string docName;
		public static string activePageCode = null;
		public static string activePlanCode = null;
		static bool onShapeAddedBreak = false;
		public static bool banOverdrawingLine = true;
		public Collection<string> whiteList = new Collection<string>() { "QueenCallout" };
		public Explorer explorer;

		public MyPage(Application _app, string _name, Explorer _explorer)
		{
			app = _app;
			docName = _name;
            explorer = _explorer;

        }

		BackgroundWorker backgroundWorker;

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
			activePageCode = getPageCode(page);
			if (activePageCode == "Plan")
			{
				if(activePlanCode == null)
				{
					activePlanCode = GetActivePlanCodeOnFirstOpen(page);
                }
                MyRibbonTracer.RibbonReload(true);
			}
			else
			{
				MyRibbonTracer.RibbonReload(false);
			}

            /*
  	            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
  	            string folderPath = System.IO.Path.Combine(docPath, "QueenScriptPictures");
  	            string fileName = page.Name + ".jpg";
  	            string fullPath = System.IO.Path.Combine(folderPath, fileName);
  	
  	            // Создаем папку если не существует
  	            if (!Directory.Exists(folderPath))
  	            {
  	                Directory.CreateDirectory(folderPath);
  	            }
  	
  	            Debug.WriteLine(fullPath);
  	            page.Export(fullPath);
  	            */

        }


        public  void onRibbonTracerBtn(IRibbonControl control)
		{
			var visioApp = Globals.ThisAddIn.Application;
			var activePage = visioApp.ActivePage;
			string pageCode = getPageCode(activePage);
			if(control.Id.Contains("btn"))
				onButtonPressed(activePage, control.Id.Substring(3));

        }

        public  void onButtonPressed(Page page, string buttonId)
		{
			if (buttonId == "AutoConnect")
			{
				if (activePlanCode != "Plan" && activePlanCode != "All")
					onTracerBtnPressed(page, activePlanCode);
			}
			else if (buttonId == "Reload")
			{
				onReloadBtnPressed(page);
			}
			else if (buttonId == "CreatePlan")
			{
				onCreatePlanPressed(page);
			}
			// Проверка по второму символу, осторожно
			else if (buttonId[1] == 'x' || buttonId == "Plan" || buttonId == "All" || buttonId.Contains("Other"))
			{
				activePlanCode = buttonId;
				onLayersBtnPressed(page);
			}
			else if (buttonId == "Lock" && GetActivePageCode() != "Plan")
			{
				foreach (Visio.Shape shape in page.Shapes)
				{
					if (Checker.isLine(shape))
					{
						Tools.CellFormulaSet(shape, "ConFixedCode", "2");
					}
				}
				setRedSquareOnPage(page, true);
				lockAllLayers(page);
			}
			else if (buttonId == "Unlock" && GetActivePageCode() != "Plan")
			{
				unlockAllLayers(page);
				setRedSquareOnPage(page, false);
			}
			else if (buttonId == "UpdatePage" && GetActivePageCode() == "planAuto")
			{
				redrawPageAuto(page);
			}
			else if (buttonId == "DevicesCheck" && GetActivePageCode() == "Plan")
			{
				Checker.CheckDevicesInPlan(page); // Не верно 
			}
			else if (buttonId == "CopyAll")
			{
				try
				{
					// ЭТО ВРЕМЯНКА, КОСТЫЛЬ БЛЯТЬ
					foreach (Visio.Shape shape in page.Shapes)
					{
						if (Checker.isLine(shape))
						{
							Tools.CellFormulaSet(shape, "ConFixedCode", "2");
						}
					}
					unlockAllLayers(page);
					Selection selection = page.Application.ActiveWindow.Selection;
					selection.SelectAll();
					selection.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | Visio.VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
					lockAllLayers(page);
					selection.DeselectAll();
				}
				catch
				{


				}
			}
			else if (buttonId == "PasteAll")
			{
				// Эксперимент
				// Смотрим, если не было ни одного "Динамическая соединительная линия.". То если они появятся, мы всех их удалим.
				/*
				int masterLineCounts = 0;
				foreach(Master master in page.Document.Masters)
					if (master.Name.Contains("Динамическая соединительная линия."))
						masterLineCounts++;
                */
				try
				{
					page.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors | Visio.VisCutCopyPasteCodes.visCopyPasteDontAddToContainers);
					setRedSquareOnPage(page, false);
					Clipboard.Clear();
					Selection selection = page.Application.ActiveWindow.Selection;
					selection.SelectAll();
					selection.ConvertToGroup();
				}
				catch
				{


				}
				/*
				if (masterLineCounts == 0) {
					bool buffer = true;
					while (buffer) {
						buffer = false;
                        foreach (Master master in page.Document.Masters)
							if (master.Name.Contains("Динамическая соединительная линия."))
							{
								buffer = true;
                                master.Delete();
							}
					}
				}
				*/
			}
			// В буфер обмена вставляем строки 
			else if (buttonId == "GetLineData")
			{
				//MessageBox.Show(CableSchedule.Generate(page),"Test",MessageBoxButtons.OK);
				string str = CableSchedule.Generate(page);
				Clipboard.SetText(str);
				Debug.WriteLine(str);

			}
			else if (buttonId == "SetHyperLinks")
			{
				SetHyperLinks(page);
            }
			else if (buttonId == "LookDevices")
			{
                explorer.LookDevices(page);
            }
			else if(buttonId == "LookDevicesOnPlan")
			{
				// Получаем соединения с плана находясь на странице устройства 
				LookDevicesOnPlan(page);
            }
        }

		public void LookDevicesOnPlan(Page page)
		{
			
			string [] pageNameCodeArray = extractGValues(page.Name);
			// Проверяем что это вообще девайс и получаем его G код
			if(pageNameCodeArray.Length == 0)
			{
                MessageBox.Show("Алгоритм не определил сигнатуру устройства","Страница не распознана",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
			}

            Visio.Page targetPage = null;
            string result = "Designation;From;Way;To;Type;Voltage;Length;Note\r\n";
            string pageNameCode = pageNameCodeArray[0].Replace("G", "");

            // Сначала ищем и получаем план в котором будет находится наш номер
            foreach (Page searchPage in page.Document.Pages)
			{
				if(Tools.CellExistsCheck(searchPage, "User.pageCode"))
				{
					// Нашли план
					if (Tools.CellFormulaGet(searchPage, "User.pageCode") == "Plan")
					{
                        // Ищем объект девайса
                        foreach (Visio.Shape shape in searchPage.Shapes)
                        {
                            if (!shape.Name.Contains("Device")) continue;

                            // Если существует 
                            if (Tools.CellExistsCheck(shape, "Prop.Number"))
                            {
                                string nameValue = shape.CellsU["Prop.Number"].FormulaU.Replace("\"", "");
								// Нашли совпадение
                                if(nameValue == pageNameCode)
								{
                                    targetPage = searchPage;
									Debug.WriteLine("НАШЛИ nameValue == pageNameCode: " + nameValue == pageNameCode);
                                    break;
								}
                            }
                        }

						// Получили страницу плана, теперь получим все соединения 
						if(targetPage != null)
						{
							string text = CableSchedule.Generate(targetPage);
							
							// Вытщаим из таблицы только строки с нашим девайсом
							foreach (string col in text.Split('\n'))
							{
								if (col.Contains("G" + pageNameCode+";") || col.Contains("G" + pageNameCode + " "))
								{
									result += col + '\n';
                                }
							}
							targetPage = null;
							
                        }

                    }
				}

			}
            // Вот тут надо вставить прям на страничку текст
            if (result.Split('\n').Length > 1)
            {

                double pageWidth = page.PageSheet.CellsU["PageWidth"].ResultIU;
                double pageHeight = page.PageSheet.CellsU["PageHeight"].ResultIU;
                double offset = 2.0 * 0.0393701;
                double left = -offset;
                double right = pageWidth + offset;
                double bottom = -offset;
                double top = pageHeight + offset;

                string table = CreateSimpleTable(result);
                Visio.Shape borderText = page.DrawRectangle(right, top, right * 2, bottom);

                borderText.Text  = "Таблица создана: " + DateTime.Today.ToString("d") + '\n' + table + '\n' +"Всего: "+ (result.Split('\n').Length-2);
                borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowFill, (short)Visio.VisCellIndices.visFillPattern].FormulaU = "0";
                borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowLine, (short)Visio.VisCellIndices.visLinePattern].FormulaU = "0";
                borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionCharacter, (short)Visio.VisRowIndices.visRowCharacter, (short)Visio.VisCellIndices.visCharacterSize].FormulaU = "9 pt";
                borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionParagraph, (short)Visio.VisRowIndices.visRowParagraph, (short)Visio.VisCellIndices.visHorzAlign].FormulaU = "0";
				borderText.CellsSRC[(short)Visio.VisSectionIndices.visSectionCharacter,(short)Visio.VisRowIndices.visRowCharacter,(short)Visio.VisCellIndices.visCharacterFont].FormulaU = "FONT(\"Cascadia Code\")";
                //MessageBox.Show(result, "Всего кабелей: " + result.Split('\n').Length, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show("Не получилось найти объект: " + pageNameCode, "Объект не найден", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }
		// Генератор таблички из csv от ИИ
        public string CreateSimpleTable(string csvData)
        {
            string[] lines = csvData.Split('\n');
            if (lines.Length == 0) return string.Empty;

            string[] headers = lines[0].Split(';');
            int[] columnWidths = new int[headers.Length];

            // Определяем ширину колонок
            for (int i = 0; i < headers.Length; i++)
            {
                columnWidths[i] = headers[i].Length;
            }

            for (int row = 1; row < lines.Length; row++)
            {
                string[] columns = lines[row].Split(';');
                for (int col = 0; col < columns.Length && col < columnWidths.Length; col++)
                {
                    columnWidths[col] = Math.Max(columnWidths[col], columns[col].Trim().Length);
                }
            }

            StringBuilder table = new StringBuilder();

            // Заголовок с разделителем
            table.Append("┌");
            for (int i = 0; i < columnWidths.Length; i++)
            {
                table.Append(new string('─', columnWidths[i] + 2));
                if (i < columnWidths.Length - 1) table.Append("┬");
            }
            table.Append("┐\n");

            table.Append("│");
            for (int i = 0; i < headers.Length; i++)
            {
                table.Append($" {headers[i].PadRight(columnWidths[i])} │");
            }
            table.Append("\n");

            // Разделитель
            table.Append("├");
            for (int i = 0; i < columnWidths.Length; i++)
            {
                table.Append(new string('─', columnWidths[i] + 2));
                if (i < columnWidths.Length - 1) table.Append("┼");
            }
            table.Append("┤\n");

            // Данные
            for (int row = 1; row < lines.Length; row++)
            {
                string[] columns = lines[row].Split(';');
                table.Append("│");

                for (int col = 0; col < columnWidths.Length; col++)
                {
                    string value = col < columns.Length ? columns[col].Trim() : "";
                    table.Append($" {value.PadRight(columnWidths[col])} │");
                }
                table.Append("\n");
            }

            // Нижняя граница
            table.Append("└");
            for (int i = 0; i < columnWidths.Length; i++)
            {
                table.Append(new string('─', columnWidths[i] + 2));
                if (i < columnWidths.Length - 1) table.Append("┴");
            }
            table.Append("┘");

            return table.ToString();
        }
        private void SetHyperLinks(Page page)
		{

            // id and array 
            List<KeyValuePair<Visio.Page, string[]>> pagesPair = new List<KeyValuePair<Visio.Page, string[]>>();
            foreach (Page _page in page.Application.ActiveDocument.Pages)
            {
                // Если активная страница - пропускаем
                if (_page?.Name == page.Name) continue;

                // ЧЕК ПО УСТРОЙСТВАМ
                Regex regex = new Regex(@"^G\d");
                if (!regex.IsMatch(_page.Name)) continue;
                //string gCode = page.Name.Split(' ').First();
                // Номера 
                pagesPair.Add(new KeyValuePair<Visio.Page, string[]>(_page, extractGValues(_page.Name)));
                Debug.WriteLine(string.Join(",", extractGValues(_page.Name)));

            }
            Debug.WriteLine("Получили");

            // Теперь пройдемся по девайсам 
            foreach (Visio.Shape shape in page.Shapes)
            {
                if (!shape.Name.Contains("Device")) continue;

                // Если существует 
                if (shape.CellExists["Prop.Number", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0)
                {
                    string nameValue = shape.CellsU["Prop.Number"].FormulaU.Replace("\"", "");
                    Debug.WriteLine("девайс " + nameValue);

                    // Если есть совпадение в pagesPair values
                    foreach (KeyValuePair<Visio.Page, string[]> keyvalue in pagesPair)
                    {
                        if (keyvalue.Value.Contains("G" + nameValue))
                        {
                            shape.AddHyperlink().SubAddress = keyvalue.Key.NameU;
                            Debug.WriteLine(shape.Name + ": '" + keyvalue.Key.Name + "'");
                            break;
                        }
                    }


                }
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
        private static string GetActivePlanCodeOnFirstOpen(Page page)
		{
			foreach (Layer layer in page.Layers)
			{
                string celIndex = layer.Index == 0 ? "" : '[' + "" + (layer.Index) + ']';
                if(page.PageSheet.Cells[$"Layers.Active" + celIndex].Formula == "1")
				{
					return layer.Name;
				}

			}
            return null;
        }

        private static void redrawPageAuto(Page page)
		{
			string pageType = page.PageSheet.CellsU["User.pageType"].Formula;
			string id = page.PageSheet.CellsU["User.id"].Formula;

			Page sPage = searchPlan(page, id);
			clearAlShapesInPage(page);

			if (sPage != null)
				copyPasteInLayer(sPage, page, GetOrCreateLayer(sPage, pageType));
		}

		private static Page searchPlan(Page _page, string id)
		{
			foreach (Page page in _page.Document.Pages)
			{

				if (page?.PageSheet.CellExists["User.pageCode", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0)
				{

					if (page.PageSheet.CellsU["User.pageCode"].Formula == "\"Plan\"")
					{
						if (page.PageSheet.CellsU["User.id"].Formula == id)
						{
							return page;
						}
					}
				}
			}
			return null;
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

        public static void OnBeforeDocumentSave(Page page)
        {
			// ПРОВЕРИТ ЧТО НЕ ПЛАН И ЕГО ПРОИЗВОДЫНЕ
            if (GetActivePageCode() == "Plan")
                return;
            setRedSquareOnPage(page, true);
            lockAllLayers(page);
        }


        // Блокировка слоя, пермещение всех фигур на отдельный слой, Переделать!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        private static void setRedSquareOnPage(Page page, bool l)
		{
			// Создаем
			if (l)
			{
				// Если уже есть такой
				foreach (Visio.Layer layer in page.Layers)
				{
					if (layer.NameU.StartsWith("RedLayer"))
					{
						return;
					}
					Marshal.ReleaseComObject(layer);

				}

				// Слой с красным квадратиком
				Visio.Layer redLayer = GetOrCreateLayer(page, "RedLayer");
				// Основной слой
				Visio.Layer mainLayer = GetOrCreateLayer(page, "Main");
				// Добавляем в основной слой
				foreach (Visio.Shape shape in page.Shapes)
				{
					try
					{
						mainLayer.Add(shape, 1);
						Marshal.ReleaseComObject(shape);
					}
					catch (Exception)
					{
					}
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
				Marshal.ReleaseComObject(borderShape);
				Marshal.ReleaseComObject(redLayer);
				Marshal.ReleaseComObject(mainLayer);
				Marshal.ReleaseComObject(borderText);

			}
			// Удаляем
			else
			{
				foreach (Visio.Layer layer in page.Layers)
				{
					if (layer.NameU.StartsWith("RedLayer"))
					{
						layer.Delete(1);
						return;
					}
					Marshal.ReleaseComObject(layer);
				}
			}
		}

		private static Visio.Layer GetOrCreateLayer(Visio.Page page, string layerName)
		{
			// Пытаемся найти существующий слой
			foreach (Visio.Layer layer in page.Layers)
			{
				if (layer.Name == layerName)
				{
					Debug.WriteLine("Найдено GetOrCreateLayer: " + layer.Name);
					return layer;
				}
			}

			// Создаем новый слой если не найден
			Visio.Layer newLayer = page.Layers.Add(layerName);
			return newLayer;
		}

		private static void onTracerBtnPressed(Page page, string activePlanCode)
		{
			autoConnect(activePlanCode, page);
		}
		private static void onReloadBtnPressed(Page page)
		{
			createNewReloadPages(page);
		}
		// Одна из кнопок связанная с линиями (Px, Cx, Plan, All, ...)
		private static void onLayersBtnPressed(Page page)
		{
			foreach (Layer layer in page.Layers)
			{
				if (activePlanCode == "All")
				{
					layerOptions(layer, 1, 1, 0, 1, 0, 0);
					continue;
				}
				// Обнуляем слой, следующими условиями все настроим
				layerOptions(layer, 0, 0, 0, 1, 0, 0);
				// Px, Cx, ...
				if (layer.Name == activePlanCode)
					layerOptions(layer, 1, 0, 1, 0, 1, 1);
				// Мы не на плане, но план должен быть отображен
				if (layer.Name == "Plan" && activePlanCode != "Plan")
					layerOptions(layer, 1, 1, 0, 1, 1, 1);
				// Если мы на плане
				else if (layer.Name == "Plan" && activePlanCode == "Plan")
					layerOptions(layer, 1, 1, 1, 0, 1, 1);
				// Слой с линиями просто скрываем, он не нужен
				if (layer.Name == "Соединительная линия")
					layerOptions(layer, 0, 0, 0, 0, 0, 0);
                if (layer.Name == "Develop")
                    layerOptions(layer, 1, 0, 0, 0, 0, 0);
                Marshal.ReleaseComObject(layer);
			}
		}

		// Настройка параметров у слоя
		// Visible, Print, Active, Lock, Snap, Glue
		private static void layerOptions(Layer layer, int v, int p, int a, int l, int s, int g)
		{
			string celIndex = layer.Index == 0 ? "" : '[' + "" + (layer.Index) + ']';
			Page page = layer.Page;
			// Видимость
			page.PageSheet.Cells[$"Layers.Visible" + celIndex].Formula = v.ToString();
			// Печать
			page.PageSheet.Cells[$"Layers.Print" + celIndex].Formula = p.ToString();
			// Активность
			page.PageSheet.Cells[$"Layers.Active" + celIndex].Formula = a.ToString();
			// Блок
			page.PageSheet.Cells[$"Layers.Locked" + celIndex].Formula = l.ToString();
			// Привязка (магнитизм)
			page.PageSheet.Cells[$"Layers.Snap" + celIndex].Formula = s.ToString();
			// Соединение 
			page.PageSheet.Cells[$"Layers.Glue" + celIndex].Formula = g.ToString();
		}
		private static void onCreatePlanPressed(Page page)
		{
			if (GetActivePageCode() == "Plan")
			{
				MessageBox.Show("План уже создан", "План", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			Selection selection = page.Application.ActiveWindow.Selection;
            if (page.Shapes.Count > 0)
			{
				MessageBox.Show("На странице обнаружены объекты, они будут перенесены на план автоматически", "Найдены объекты", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				selection.SelectAll();
            }
			short sectionUser = (short)VisSectionIndices.visSectionUser;
			int planCount = getAllPlanCount(page);
			int id = planCount;

			// Добавляем секцию юзер
			page.PageSheet.AddSection(sectionUser);

			short row;
			// Plan
			row = (short)page.PageSheet.AddNamedRow(sectionUser, "pageCode", (short)VisRowTags.visTagDefault);
			page.PageSheet.CellsSRC[sectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"\"{"Plan"}\"";
			// id
			row = (short)page.PageSheet.AddNamedRow(sectionUser, "id", (short)VisRowTags.visTagDefault);
			page.PageSheet.CellsSRC[sectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"\"{id}\"";

            // scale
            page.PageSheet.AddSection((short)VisSectionIndices.visSectionProp);
            row = (short)page.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionProp, "Scale", (short)VisRowTags.visTagDefault);
            page.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionProp, row, (short)VisCellIndices.visUserValue].FormulaU = $"1";
            page.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionProp, row, (short)VisCellIndices.visCustPropsLabel].FormulaU = $"\"Размер элементов\"";


            setPrintOffOnLayers(page);

			foreach (var item in WireService.wires)
			{
				page.Layers.Add(item.name);
				if (item.name != "Plan")
					setLayerSheetPrint(page, item.name, 0);
			}
			MyRibbonTracer.RibbonReload(true);
			activePageCode = "Plan";

			// Переносим все на PLAN
			foreach (Visio.Shape shape in selection)
			{
                page.Layers["Plan"].Add(shape, 0);
            }
        }

		public static void createNewReloadPages(Page page)
		{
			// Защита от дурака
			if (GetActivePageCode() != "Plan")
				return;


            Thread th = new Thread(() =>
			{
				// Костыль
				onShapeAddedBreak = true;

				Selection selectionPlan = page.Application.ActiveWindow.Selection;
				selectionPlan = page.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper, "Plan");

				// Удаляем старое и создаем подложку с устройствами 
				Page backPage = GetOrCreateBackgroundPlan(page, page.PageSheet.CellsU["User.id"].Formula);
				// На подложку вставляем устройства
				try
				{
					selectionPlan.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
					backPage.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
					Clipboard.Clear();
				}
				catch
				{
				}

                Collection<Visio.Shape> backGroundShapes = new Collection<Visio.Shape>();
                //Selection backGroundShapes = page.Application.ActiveWindow.Selection;

                foreach (Visio.Shape shape in page.Shapes)
				{
					if (shape.Name.Contains("Device") || shape.Name.Contains("Shield") || shape.Name.Contains("Recorder") || shape.Name.Contains("WireName"))
					{
                        backGroundShapes.Add(shape);
                    }
				}

                // Эта переменная растетб что бы странички шли друг за другом ровно, иначе если страница была пропущена, индекс пойдет дальше и будет перескок
                int minusIndex = 0;
                //Collection<Page> pagestest = new Collection<Page>();

                foreach (Layer layer in page.Layers)
				{
                    // Детектор сделали ли мы что то на слое чи не
                    // Upd ориентир по прозрачности 1%

                    // Проверяем все слои что есть в списке
                    if (WireService.ThatIsTraccer(layer.Name))
					{
                        // Тут было условие по кол ву кабелей на слое 
                        //Создаем селекциию, на всякий очищаем находящейся там
                        Selection selection = page.Application.ActiveWindow.Selection;
                        // Выделяем все фигуры на слое и слой плана, добавляем в селкцию
                        selection = page.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer, Visio.VisSelectMode.visSelModeSkipSuper, layer.Name);

						if(selection.Count == 0)
						{
							minusIndex++;
                        }
						// Если ничего нет на слое, то скип
						else
						{
							// Поиск существующих
							Page defNewPage = searchAutoPage(page, page.PageSheet.CellsU["User.id"].Formula, layer.Name);
							// Страница уже существует, удаляем
							if (defNewPage != null)
								defNewPage.Delete(2);

							short row;
							// Создаем страницу
							defNewPage = page.Document.Pages.Add();
							// Делаем что бы следующей шла и минусуем если слой пропущен
							defNewPage.Index = (short)(page.Index + WireService.wires.FindIndex(w => w.name == layer.Name) - minusIndex);
							// Имя страницы
							Wire wire = WireService.wires.FirstOrDefault(w => w.name == layer.Name);
							string  _pagename1 = page.Name.Replace("Plan","").Replace("-","").Trim();
                            defNewPage.Name = "Plan " + layer.Name + " - " + _pagename1 +" "+ wire.comment;
							// Даем код страницы ("planAuto")
							row = (short)defNewPage.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionUser, "pageCode", (short)VisRowTags.visTagDefault);
							defNewPage.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"\"{"planAuto"}\"";
							// Даем тип страницы (Ax, Px, ...)
							row = (short)defNewPage.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionUser, "pageType", (short)VisRowTags.visTagDefault);
							defNewPage.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"\"{layer.Name}\"";
							// Даем Id (0,1,2 ...)
							row = (short)defNewPage.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionUser, "id", (short)VisRowTags.visTagDefault);
							defNewPage.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"{page.PageSheet.CellsU["User.id"].Formula}";

							defNewPage.BackPage = backPage;

							// Запоминаем был ли он visible чи не
							string celIndex = layer.Index == 0 ? "" : '[' + "" + (layer.Index) + ']';
							string memoryVisibleCell = page.PageSheet.CellsU[$"Layers.Visible" + celIndex].Formula;
							page.PageSheet.CellsU[$"Layers.Visible" + celIndex].Formula = "1";


							// Если есть что выделять (на случай пустых страниц)
							if (selection.Count > 0)
							{
                                // Заранее у нас просчитана колекция с девайсами и щитами, что бы они наложились поверх
                                foreach (Visio.Shape shape in backGroundShapes)
                                {
                                    selection.Select(shape, 2);
                                }
                                try
								{

                                    selection.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
									defNewPage.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
									Clipboard.Clear();

									//pagestest.Add(defNewPage);

                                    //selectionDevices.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
                                    //defNewPage.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
                                    //Clipboard.Clear();
                                }
								catch
								{
								}

								// Тут ищем вдруг есть WireName Объект и даем ему название кабеля
								foreach (Visio.Shape shape in defNewPage.Shapes)
								{
									if (shape.Name.Contains("WireName"))
									{
										// Нашли									
                                        shape.Text = wire.comment + " " + wire.defaultCable + " " + wire.voltage; 
                                        shape.CellsU["Char.Color"].FormulaU = wire.color;

										// Если мы на странице света, ищем RGB что бы указать их как UTP
										if(layer.Name == "Lx")
                                        {
											List<string> lxNames = new List<string>();
                                            foreach (Visio.Shape shape1 in defNewPage.Shapes)
                                            {
                                                if (shape1.Name.Contains("Light"))
                                                {
                                                    if (shape1.CellExists["Prop.Type", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0)
													{
														// Нашли RGB
														if(shape1.CellsU["Prop.Type"].FormulaU.Replace("\"", "") == "INDEX(2,Prop.Type.Format)")
														{
															lxNames.Add("L" + shape1.CellsU["Prop.Number"].FormulaU.Replace("\"", ""));
                                                        }
                                                    }
                                                }
                                            }
											// Есть хоть одно
											if(lxNames.Count != 0)
											{
												shape.Text += '\n';
                                                shape.Text += string.Join(", ", lxNames);
												shape.Text += " - UTP Cat 5E";
                                            }
                                        }
                                        break;
									}
								}

							}
                            selection.DeselectAll();


                            // Видимый для печати 
                            setPrintOnOnLayers(defNewPage);
							// Блокируем
							lockAllLayers(defNewPage);
							// Возращаем значение которое запомнили 
							page.PageSheet.CellsU[$"Layers.Visible" + celIndex].Formula = memoryVisibleCell;



						}

					}

					// Очищаем ком
					Marshal.ReleaseComObject(layer);

				}
				/*
				foreach (Page item in pagestest)
				{
                    selectionDevices.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
                    item.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate | Visio.VisCutCopyPasteCodes.visCopyPasteNoHealConnectors);
                    Clipboard.Clear();
                    Marshal.ReleaseComObject(item);

                }
				*/
                Marshal.ReleaseComObject(backPage);
				onShapeAddedBreak = false;
			}); th.SetApartmentState(ApartmentState.STA); th.IsBackground = true; th.Start();
			

		}


		// Нам нужен метод, работающий на плане, во время трассровки, назначит линиям User.code 
		// Еще нужны настройки автоматические


		// Функция проходит по всем страничкам (задом наперед) ищет фоновую страницу которая соотвествует айди плана, очищает ее и возращает. Либо если ее нет, создает новую и возвращает
		static private Page GetOrCreateBackgroundPlan(Page page, string id)
		{
			
			// Сначала ищем
			foreach (Page _page in page.Document.Pages.Cast<Page>().Reverse())
			{
				// Если это фоновая
				if (_page.Background != 0)
				{
					Debug.WriteLine(_page.Name);
					if (_page?.PageSheet != null && (_page?.PageSheet.CellExists["User.pageCode", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0) &&
						_page?.PageSheet != null && (_page?.PageSheet.CellExists["User.id", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0))
					{
						if (_page.PageSheet.CellsU["User.pageCode"].Formula == '"' + "PlanBackGround" + '"' &&
							_page.PageSheet.CellsU["User.id"].Formula == id )
						{
							// Очищаем и возвращаем
							Selection sel = _page.CreateSelection(Visio.VisSelectionTypes.visSelTypeAll, Visio.VisSelectMode.visSelModeSkipSuper);
							// НЕ УДАЕТСЯ ИЗ ЗА БЛОКИРОВАННОГО 
							sel.Delete();
							return _page;
						}

					}

				}
			}
			// Если страничка не найдена 
			// Создаем страницу
			Page newPage = page.Document.Pages.Add();
			short row;
			// Делаем что она была в конце
			newPage.Index = (short)(page.Document.Pages.Count - 1);
			// Делаем фоном
			newPage.Background = 1;
			// Имя страницы
			newPage.Name = "DEVELOP " + page.Name + " backround";
			// Даем код страницы ("planAuto")
			row = (short)newPage.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionUser, "pageCode", (short)VisRowTags.visTagDefault);
			newPage.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = $"\"{"PlanBackGround"}\"";
			// Даем Id (0,1,2 ...)
			row = (short)newPage.PageSheet.AddNamedRow((short)VisSectionIndices.visSectionUser, "id", (short)VisRowTags.visTagDefault);
			newPage.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionUser, row, (short)VisCellIndices.visUserValue].FormulaU = id;
			return newPage;
		}

		// page == plan page!!!!!!!!!!!
		private static void copyPasteInLayer(Page page, Page defNewPage, Layer layer)
		{


		}
	

		// Удаляет все что на странице
		private static void clearAlShapesInPage(Page page)
		{
			unlockAllLayers(page);
			while (page.Shapes.Count > 0)
			{
				page.Shapes[1].CellsU["LockDelete"].Formula = "0";
				page.Shapes[1].Delete();
			}
		}

		private static void unlockAllLayers(Page page)
		{
			int i = 0;
			foreach (Layer item in page.Layers)
			{
				string celIndex = i == 0 ? "" : '[' + "" + (i + 1) + ']';
				page.PageSheet.Cells[$"Layers.Locked" + celIndex].Formula = "0";
				page.PageSheet.Cells[$"Layers.Active" + celIndex].Formula = "0";

				i++;
				Marshal.ReleaseComObject(item);
			}
		}

		private static void lockAllLayers(Page page)
		{
			int i = 0;
			foreach (Layer item in page.Layers)
			{
				string celIndex = i == 0 ? "" : '[' + "" + (i + 1) + ']';
				page.PageSheet.Cells[$"Layers.Locked" + celIndex].Formula = "1";
				page.PageSheet.Cells[$"Layers.Active" + celIndex].Formula = "0";

				i++;
				Marshal.ReleaseComObject(item);
			}

		}

		private static void setPrintOffOnLayers(Page page)
		{
			int i = 0;
			foreach (Layer item in page.Layers)
			{
				string celIndex = i == 0 ? "" : '[' + "" + (i + 1) + ']';
				if (item.Name == "Plan")
				{
					i++;
					continue;
				}
				page.PageSheet.Cells[$"Layers.Print" + celIndex].Formula = "0";

				i++;
				Marshal.ReleaseComObject(item);
			}
		}
		private static void setPrintOnOnLayers(Page page)
		{
			int i = 0;
			foreach (Layer item in page.Layers)
			{
				string celIndex = i == 0 ? "" : '[' + "" + (i + 1) + ']';
				if (item.Name == "Plan")
				{
					i++;
					continue;
				}
				page.PageSheet.Cells[$"Layers.Print" + celIndex].Formula = "1";

				i++;
				Marshal.ReleaseComObject(item);
			}
		}

		private static void setLayerSheetPrint(Page page, string row, short value)
		{
			int i = 0;
			foreach (Layer item in page.Layers)
			{
				string celIndex = i == 0 ? "" : '[' + "" + (i + 1) + ']';
				if (item.Name == row)
					page.PageSheet.Cells[$"Layers.Print" + celIndex].Formula = value + "";

				i++;
				Marshal.ReleaseComObject(item);
			}
			//Marshal.ReleaseComObject(page);
			//short num = page.Layers.ItemU['"' +row+'"'].Index;
			//page.PageSheet.CellsSRC[(short)VisSectionIndices.visSectionLayer, (short)VisCellIndices.visLayerPrint, num].FormulaU = value + "" ;
		}


		// Поиск уже автоматически созданных страниц в документе. Автотрассировка.
		public static Page searchAutoPage(Page page, string id, string wireCode)
		{
			foreach (Page item in page.Document.Pages)
			{

				if (item?.PageSheet != null && (item?.PageSheet.CellExists["User.pageType", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0) &&
					item?.PageSheet != null && (item?.PageSheet.CellExists["User.id", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0) &&
					item?.PageSheet != null && (item?.PageSheet.CellExists["User.pageCode", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0))
				{
					if (item.PageSheet.CellsU["User.pageType"].Formula == '"' + wireCode + '"' &&
				   item.PageSheet.CellsU["User.id"].Formula == id &&
				   item.PageSheet.CellsU["User.pageCode"].Formula == "\"planAuto\"")
					{
						return item;
					}
				}

			}
			return null;
		}

		public static int getAllPlanCount(Page page)
		{
			int i = 0;
			foreach (Page item in page.Document.Pages)
			{
				if (item?.PageSheet != null && (item?.PageSheet.CellExists["User.pageCode", (short)Visio.VisExistsFlags.visExistsAnywhere] != 0))
				{
					if (item.PageSheet.CellsU["User.pageCode"].Formula == "\"Plan\"")
					{
						i++;
					}
				}
			}
			return i;
		}

		private static void autoConnect(string code, Page page)
		{
			Visio.Shape shield = null;
			List<Visio.Shape> devices = new List<Visio.Shape>();

            if (!WireService.wires.Any(w => w.name == activePlanCode && w.isWire))
            {
                MessageBox.Show("Ошибка слоя","Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }

            var selection = page.CreateSelection(Visio.VisSelectionTypes.visSelTypeByLayer,Visio.VisSelectMode.visSelModeSkipSuper,
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
		public static void CreateAndGlueConnector(Page page, Visio.Shape fromShape, Visio.Shape toShape, string layerName)
		{
			// Костыль
			rebuildBrake = true;
            Visio.Shape connector = CreateConnector(page);
			Debug.WriteLine(connector.Name);
            GlueConnectorToShapes(connector, fromShape, toShape);
            rebuildBrake = false;
            rebuildShape(connector, true);

            //Marshal.ReleaseComObject(connector);
        }
        // Костыль
        public static bool rebuildBrake = false;

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

		private static bool checkConnetctedLinesInDevice(Page page, Visio.Shape shape)
		{
			foreach (Visio.Shape connector in page.Shapes.Cast<Visio.Shape>().Where(s => Checker.isLine(s)))
			{
				if (isShapeOnLayer(connector, activePlanCode))
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

		private static bool isShapeOnLayer(Visio.Shape shape, string layerName)
		{
			if (shape.Layer[2].Name == layerName)
				return true;

			return false;
		}


		private static void newLayerDetect(Visio.Page page)
		{
			if (page.Layers.Count == WireService.wires.Count)
				return;

			foreach (Layer layer in page.Layers)
			{
				if (layer.Name == "" || layer.Name == "Соединительная линия")
					continue;

				bool result = false;

				foreach (Wire wire in WireService.wires)
				{

					if (wire.name == layer.Name)
					{
						result = true;
					}
				}
				if (!result)
				{
					Debug.WriteLine("Найдено лишнее: " + layer.Name);
					replaceLayer(page, layer.Name, activePlanCode);
				}
				Marshal.ReleaseComObject(layer);
			}

		}

		// Элементы с одного слоя перемещаем на другой слой
		private static void replaceLayer(Page page, string sourceLayerName, string targetLayerName)
		{
			// Ищем слои
			Layer target = null;
			Layer source = null;
			foreach (Layer layer in page.Layers)
			{
				if (layer.Name == targetLayerName)
				{
					target = layer;
					break;
				}
				Marshal.ReleaseComObject(layer);
			}
			foreach (Layer layer in page.Layers)
			{
				if (layer.Name == sourceLayerName)
				{
					source = layer;
					break;
				}
				Marshal.ReleaseComObject(layer);
			}

			if (target is null || source is null) return;


			foreach (Visio.Shape shape in page.Shapes)
			{
				for (short i = 1; i < shape.LayerCount + 1; i++)
				{
					if (shape.Layer[i].Name == sourceLayerName)
					{
						target.Add(shape, 0);
					}
				}
				Marshal.ReleaseComObject(shape);
			}
			layerOptions(source, 0, 0, 0, 0, 0, 0);
			source.Delete(1);
		}


		public void onShapeChanged(Visio.Shape shape)
		{
			Debug.WriteLine("Shape changed: " + shape.Name);
		}

		// Получение кода страницы
		private static string getPageCode(Page page)
		{
			// Существует ли у страницы pageCode
            if (Tools.CellExistsCheck(page, "User.pageCode"))
			{
				Debug.WriteLine(Tools.CellFormulaGet(page, "User.pageCode"));
				return Tools.CellFormulaGet(page, "User.pageCode");
			}
			return null;
		}

		public static string GetActivePageCode()
		{
			return activePageCode;
		}

		public void onShapeAdded(Visio.Shape shape)
		{
			if (GetActivePageCode() == "Plan") // ??
			{
                // Костыль
                if (!rebuildBrake)
					rebuildShape(shape, true);
				// Чекаем вдруг появился новый layer
				newLayerDetect(shape.Application.ActivePage);
			}

			if (banOverdrawingLine)
			{
				if (Checker.isLine(shape))
				{
                    Tools.CellFormulaSet(shape, "ConFixedCode", "2");
                }
            }
		}

		// onUserShape - false: Значит линия создана была автоматически
		private static void rebuildShape(Visio.Shape shape, bool onUserShape = false)
		{
			// Костыль
			if (onShapeAddedBreak)
			{
				return;
			}

			if (Checker.isLine(shape) && GetActivePageCode() == "Plan") // ?? 
			{
				if (activePlanCode == null)
				{
					shape.Delete();
					return;
				}

				// Если она ничкему не присоеденена
				if (shape.CellsU["BegTrigger"].FormulaU.Contains("Dynamic connector.") ||
				   shape.CellsU["EndTrigger"].FormulaU.Contains("Dynamic connector."))
				{
					shape.Delete();
					return;
				}
				// На плане не даем размещать линии
				if (activePlanCode == "Plan" /*|| (activePlanCode != "Plan" && shape.Name.Contains("Device"))*/)
				{
					shape.Delete();
					return;
				}
				// Если мы на трасировке,добавляем юзершейпы на линии
				else if (WireService.wires.Any(w => w.name == activePlanCode && w.isWire)) //(activePlanCode[1] == 'x' || activePlanCode.Contains("Other"))
				{
					// Свапнем конец если не так подключили 
                    if (shape.Connects.Count >= 2)
                    {
                        Visio.Shape connectedShapeFrom = shape.Connects[1].ToSheet;
                        Visio.Shape connectedShapeTo = shape.Connects[2].ToSheet;

                        if (connectedShapeTo.Name.Contains("Shield"))
                        {
							// Костыль
                            //shape.Delete();
                            //rebuildBrake = true;
                            //shape = CreateConnector(connectedShapeTo.Application.ActivePage);
                            //GlueConnectorToShapes(shape, connectedShapeFrom, connectedShapeTo);
                            //rebuildBrake = false;
                        }
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

					Visio.Shape nearestLine = FindNearestLine(shape);
					if(nearestLine != null)
						MergeLineGeometry(nearestLine, shape);

                }
            }

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
                if (Checker.isLine(shape))
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
			Debug.WriteLine(connectedShape.Name);

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

                    short row;
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
