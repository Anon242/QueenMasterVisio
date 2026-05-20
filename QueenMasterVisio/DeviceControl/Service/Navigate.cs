using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;

namespace QueenMasterVisio.DeviceControl.Service
{
    internal class Navigate
    {
        public static void NavigateToShape(Application visioApp, Shape shape)
        {
            try
            {
                if (visioApp?.ActiveDocument == null)
                    return;

                visioApp.ActiveWindow.CenterViewOnShape(shape, VisCenterViewFlags.visCenterViewDefault);
           
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка перехода: {ex.Message}", "Visio навигация",
                                                System.Windows.MessageBoxButton.OK,
                                                System.Windows.MessageBoxImage.Error);
            }
        }

        public static void NavigateToPage(Application visioApp, Page page)
        {
            try
            {
                if (visioApp?.ActiveDocument == null)
                    return;

                visioApp.ActiveWindow.Page = page;

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка перехода: {ex.Message}", "Visio навигация",
                                                System.Windows.MessageBoxButton.OK,
                                                System.Windows.MessageBoxImage.Error);
            }
        }
    }
}
