using QueenMasterVisio.Core.Helpers;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;

namespace QueenMasterVisio.Core.Managers
{
    internal class PageManager
    {
        public static void createNewReloadPages(Page page)
        {
            // Защита от дурака
            if (!page.IsPlanPage())
                return;


            Thread th = new Thread(() =>
            {
                // Костыль
                using (VisioEventSuppressor.SuppressShapeAdded())
                {

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

                    // Эта переменная растет что бы странички шли друг за другом ровно, иначе если страница была пропущена, индекс пойдет дальше и будет перескок
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

                            if (selection.Count == 0)
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
                                string _pagename1 = page.Name.Replace("Plan", "").Replace("-", "").Trim();
                                defNewPage.Name = "Plan " + layer.Name + " - " + _pagename1 + " " + wire.comment;
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
                                if (selection.Count > 1)
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
                                            Debug.WriteLine("Мы написали");
                                            // Если мы на странице света, ищем RGB что бы указать их как UTP
                                            Debug.WriteLine("layer.Name " + layer.Name);

                                            if (layer.Name == "Lx")
                                            {
                                                Debug.WriteLine("Опредилили " + layer.Name);

                                                List<string> lxNames = new List<string>();
                                                foreach (Visio.Shape shape1 in defNewPage.Shapes)
                                                {
                                                    if (shape1.Name.Contains("Light"))
                                                    {
                                                        Debug.WriteLine("СВЕТ " + shape1.Name);

                                                        if (Tools.CellExistsCheck(shape1, "Prop.Type"))
                                                        {
                                                            // Нашли RGB
                                                            Debug.WriteLine(shape1.CellsU["Prop.Type"].FormulaU);
                                                            if (Tools.CellExistsCheck(shape1, "Prop.Type"))
                                                            {
                                                                string value = Tools.CellValueGet(shape1, "Prop.Type");
                                                                if (value.Contains("RGB"))
                                                                {
                                                                    lxNames.Add("L" + shape1.CellsU["Prop.Number"].FormulaU.Replace("\"", ""));
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                // Есть хоть одно
                                                if (lxNames.Count != 0)
                                                {
                                                    shape.Text += '\n';
                                                    shape.Text += string.Join(", ", lxNames);
                                                    shape.Text += " - 4 x 1.5";
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
                }
            }); th.SetApartmentState(ApartmentState.STA); th.IsBackground = true; th.Start();


        }

    }
}
