using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Visio = Microsoft.Office.Interop.Visio;
using Shape = Microsoft.Office.Interop.Visio.Shape;
using Page = Microsoft.Office.Interop.Visio.Page;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Managers;

namespace QueenMasterVisio.Core.Services
{
    /// <summary>
    /// Класс для создания новой страницы с созданием его кабелей
    /// </summary>
    internal class CreateDeviceInPlan
    {
        public static void CreatePage(Page page, Visio.Selection selection)
        {
            
            if (selection.Count == 1)
            {
                Shape shape = selection[1];
                if (shape.NameU.Contains("Device")) 
                {
                    string deviceName = shape.GetCellResultString("Prop.number");
                    if (string.IsNullOrEmpty(deviceName))
                        return;

                    string pageName = QueenMasterVisio.RenameDialog.ShowRenameDialog("G"+ deviceName + " <имя>");

                    if (string.IsNullOrEmpty(pageName))
                        return;

                    Page newPageDevice = Managers.DocumentManager.CreateNewPage(pageName);

                    // Получаем базу данных девайса
                    string cableShedule = CableService.Generate(page);
                    CableManager cb = new CableManager(newPageDevice);
                    int i = 0;
                    foreach (string item in cableShedule.Split('\n'))
                    {
                        if(item.Contains("G" + deviceName + ";"))
                        {
                            // Designation;From;Way;To;Type;Voltage;Length;Note
                            string[] arr = item.Split(';');
                            if (arr[4].Contains("UTP"))
                                arr[4] = "UTP";
                            cb.Create(arr[4].Trim(' '), arr[2], arr[1],i++, -1);
                        }
                    }
                    // Cоздаем кабеля


                }
                // Можно еще свет
                else
                {
                    MessageBox.Show("Выделите объект девайса");
                }
            }
            else
            {
                MessageBox.Show("Ничего не выделено либо выделено больше одного объекта");
            }
        }
    }
}
