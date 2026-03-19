using Microsoft.Office.Core;
using System;
using System.Collections.Generic;

using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using QueenMasterVisio.Core.Handlers;

// TODO:  Выполните эти шаги, чтобы активировать элемент XML ленты:

// 1: Скопируйте следующий блок кода в класс ThisAddin, ThisWorkbook или ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MainLentXml();
//  }

// 2. Создайте методы обратного вызова в области "Обратные вызовы ленты" этого класса, чтобы обрабатывать действия
//    пользователя, например нажатие кнопки. Примечание: если эта лента экспортирована из конструктора ленты,
//    переместите свой код из обработчиков событий в методы обратного вызова и модифицируйте этот код, чтобы работать с
//    моделью программирования расширения ленты (RibbonX).

// 3. Назначьте атрибуты тегам элементов управления в XML-файле ленты, чтобы идентифицировать соответствующие методы обратного вызова в своем коде.  

// Дополнительные сведения можно найти в XML-документации для ленты в справке набора средств Visual Studio для Office.


namespace QueenMasterVisio.Ribbon
{
    [ComVisible(true)]
    public class MainLentXml : Office.IRibbonExtensibility
    {
        private static Office.IRibbonUI ribbon;

        public static bool ribbonVisible = false;

        RibbonCommandHandler ribbonCommandHandler;
        static string layerButtonName = "";
        public MainLentXml()
        {
            ribbonCommandHandler = new RibbonCommandHandler();
        }

        public static void RibbonReload(bool vis)
        {
            ribbonVisible = vis;
            ribbon?.Invalidate();
            ribbon?.ActivateTab("MyTab");

        }
        // Сделать отдельный класс который определить кто куда что вызывает
        public void OnButtonClickPlan(IRibbonControl control)
        {
            ribbonCommandHandler.onRibbonTracerBtnPlan(control);
        }

        public void OnButtonClickDevice(IRibbonControl control)
        {
            ribbonCommandHandler.onRibbonTracerBtnDevice(control);

        }

        public string GetLayerButtonLabel(IRibbonControl control)
        {
            string code = control.Id.Replace("btn", "");
            return (code == layerButtonName) ? (code + "\n+") : code;
        }

        public static void UpdateLayerButtons(string _layerButtonName)
        {
            layerButtonName = _layerButtonName;
        }

        public bool GetCheckboxState(IRibbonControl control)
        {
            // Возвращает true для включенного состояния по умолчанию
            return true;
        }
        /*
        public void OnCheckboxClick(IRibbonControl control, bool isPressed)
        {
            < checkBox id = ""chkToogleLine"" label = ""Убрать авто-перетрассировку линий"" getPressed = ""GetCheckboxState"" onAction = ""OnCheckboxClick"" />
        }
        */

        public bool GetTabVisible(IRibbonControl control)
        {
            if (ribbonVisible)
                return true;
            else
                return false;
        }


        #region Элементы IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("QueenMasterVisio.Ribbon.MainLentXml.xml");
        }

        #endregion

        #region Обратные вызовы ленты
        //Информацию о методах создания обратного вызова см. здесь. Дополнительные сведения о методах добавления обратного вызова см. по ссылке https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        #endregion

        #region Вспомогательные методы

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
