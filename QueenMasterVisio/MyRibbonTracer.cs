using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Visio;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;
using System.Diagnostics;

namespace QueenMasterVisio
{
    [ComVisible(true)]
    public class MyRibbonTracer : Office.IRibbonExtensibility
    {
        private static Office.IRibbonUI _ribbon;
        private  MyRibbonTracer _instance;
        public static bool ribbonVisible = false;
        MyPage myPage;

        public MyRibbonTracer()
        {
            _instance = this; 
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            Debug.WriteLine("Loaded");
            Debug.WriteLine(_ribbon);
        }

        public static void RibbonReload(bool vis)
        {
            ribbonVisible = vis;
            _ribbon?.Invalidate();
            _ribbon?.ActivateTab("MyTab");
           
        }
        public string GetCustomUI(string ribbonID) => @"<customUI onLoad=""OnRibbonLoad"" xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
        <tab id=""MyTab2"" label=""Queen мастер"">
            <group id=""MyGroup6"" label=""Создание"">
                <button id=""btnCreatePlan"" label=""Создать из этой страницы план"" size=""normal"" onAction=""OnButtonClick"" imageMso=""ZoomOnePage""/>
                <button id=""btnUpdatePage"" label=""Обновить страницу"" size=""normal"" onAction=""OnButtonClick"" imageMso=""AutoFormatChange""/>
            </group>
            <group id=""MyGroup61"" label=""Блокировка"">
                <button id=""btnLock"" label=""Заблокировать страницу"" size=""large"" onAction=""OnButtonClick"" imageMso=""ColumnActionsReadOnly""/>
                <button id=""btnUnlock"" label=""Разблокировать страницу"" size=""large"" onAction=""OnButtonClick"" imageMso=""CustomizeMySite""/>
            </group>
            <group id=""MyGroup62"" label=""Копирование"">
                <button id=""btnCopyAll"" label=""Скопировать страницу"" size=""large"" onAction=""OnButtonClick"" imageMso=""Copy""/>
                <button id=""btnPasteAll"" label=""Вставить все на страницу"" size=""large"" onAction=""OnButtonClick"" imageMso=""Paste""/>
            </group>
           <group id=""MyGroup63"" label=""Настройки"">
                <checkBox id=""chkToogleLine"" label=""Убрать авто-перетрассировку линий"" getPressed=""GetCheckboxState"" onAction=""OnCheckboxClick""/>
            </group>
           <group id=""MyGroup64"" label=""Информация"">
                <button id=""btnLookDevices"" label=""Показать устройства"" size=""large"" onAction=""OnButtonClick"" imageMso=""CustomizeMySite""/>
            </group>
        </tab>
      <tab id=""MyTab"" label=""План трассеров"" getVisible=""GetTabVisible"">
        <group id=""MyGroup2"" label=""Слои"">
            <box id=""btnGroup"" boxStyle=""horizontal"">
          <button id=""btnPlan"" label=""План"" size=""large"" onAction=""OnButtonClick"" imageMso=""DesignMode""/>
          <button id=""btnPx"" label=""Px"" size=""large"" onAction=""OnButtonClick"" imageMso=""ShapeLightningBolt""/>
          <button id=""btnEx"" label=""Ex"" size=""large"" onAction=""OnButtonClick"" imageMso=""QuickPartsInsertFromOnline""/>
          <button id=""btnRx"" label=""Rx"" size=""large"" onAction=""OnButtonClick"" imageMso=""OrganizationChartSelectAllConnectors""/>
          <button id=""btnDx"" label=""Dx"" size=""large"" onAction=""OnButtonClick"" imageMso=""AutoFormatChange""/>
          <button id=""btnCx"" label=""Cx"" size=""large"" onAction=""OnButtonClick"" imageMso=""OrganizationChartAutoLayout""/>
          <button id=""btnSx"" label=""Sx"" size=""large"" onAction=""OnButtonClick"" imageMso=""FileCheckIn""/>
          <button id=""btnYx"" label=""Yx"" size=""large"" onAction=""OnButtonClick"" imageMso=""ReviewReplyWithChanges""/>
          <button id=""btnVx"" label=""Vx"" size=""large"" onAction=""OnButtonClick"" imageMso=""RecordsSaveRecord""/>
          <button id=""btnLx"" label=""Lx"" size=""large"" onAction=""OnButtonClick"" imageMso=""_3DLightingClassic""/>
          <button id=""btnAx"" label=""Ax"" size=""large"" onAction=""OnButtonClick"" imageMso=""NewContact""/>
          <button id=""btnOther1"" label=""Разное 1"" size=""large"" onAction=""OnButtonClick"" imageMso=""RightToLeftDocument""/>
          <button id=""btnOther2"" label=""Разное 2"" size=""large"" onAction=""OnButtonClick"" imageMso=""HeaderFooterPageNumberInsert""/>
          <button id=""btnAll"" label=""Показать все"" size=""large"" onAction=""OnButtonClick"" imageMso=""CategorizeMenu""/>
 </box>
        </group>
        <group id=""MyGroup3"" label=""Инструменты"">
          <!-- Указатель -->
          <control idMso=""PointerTool"" size=""large""/>
          <!-- Линия -->
          <control idMso=""ConnectorTool"" size=""large""/>
          <!-- Прямоугольник -->
          <control idMso=""RectangleTool"" size=""large""/>
          <!-- Текст -->
          <control idMso=""TextTool"" size=""large""/>
        </group>
        <group id=""MyGroup5"" label=""Автоматическое соединение"">
          <button id=""btnAutoConnect"" label=""Трассировать на слое"" size=""large"" onAction=""OnButtonClick"" imageMso=""ArrowStyleGallery""/>
        </group>
        <group id=""MyGroup4"" label=""Обновить"">
          <button id=""btnReload"" label=""Обновить все слои для печати"" size=""large"" onAction=""OnButtonClick"" imageMso=""BuildingBlocksSaveTableOfContents""/>
        </group>
        <group id=""MyGroup7"" label=""Проверка"">
          <button id=""btnDevicesCheck"" label=""Диагностика страницы"" size=""large"" onAction=""OnButtonClick"" imageMso=""BuildingBlocksSaveTableOfContents""/>
        </group>
        <group id=""MyGroup71"" label=""Данные"">
          <button id=""btnGetLineData"" label=""Генерация кабель-журнала"" size=""large"" onAction=""OnButtonClick"" imageMso=""BuildingBlocksSaveTableOfContents""/>
          <button id=""btnSetHyperLinks"" label=""Назначить ссылки"" size=""large"" onAction=""OnButtonClick"" imageMso=""BuildingBlocksSaveTableOfContents""/>

        </group>
      </tab>

    </tabs>
  </ribbon>
</customUI>";

        public void OnButtonClick(IRibbonControl control)
        {
            //Debug.WriteLine(control.Id);
            myPage.onRibbonTracerBtn(control);
        }

        public void SetMyPage(MyPage _myPage)
        {
            myPage = _myPage;
        }
        public bool GetCheckboxState(IRibbonControl control)
        {
            // Возвращает true для включенного состояния по умолчанию
            return true;
        }
        public void OnCheckboxClick(IRibbonControl control, bool isPressed)
        {
            MyPage.banOverdrawingLine = isPressed;
        }


        public bool GetTabVisible(IRibbonControl control)
        {
            if (ribbonVisible)
                return true;
            else
                return false;
        }


    }
}
