using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.DeviceControl.Results;
using QueenMasterVisio.DeviceControl.Rules;
using QueenMasterVisio.DeviceControl.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = Microsoft.Office.Interop.Visio.Application;
using Page = Microsoft.Office.Interop.Visio.Page;

namespace QueenMasterVisio.DeviceControl
{
    /// <summary>
    /// Логика взаимодействия для DeviceControlForm.xaml
    /// </summary>
    public partial class DeviceControlForm : System.Windows.Controls.UserControl
    {
        private ICollectionView _view;
        private Application _app;
        public DeviceControlForm(Application app)
        {
            _app = app;
            InitializeComponent();
            
            //SetupFiltering();
        }
        private void SetupFiltering()
        {
            _view = CollectionViewSource.GetDefaultView(dataGrid.ItemsSource);
            _view.Filter = item => FilterPredicate(item as Models.FormItem);

            _view.SortDescriptions.Clear();
            _view.SortDescriptions.Add(new SortDescription("Code", ListSortDirection.Descending));
        }

        private void FilterToggle_Checked(object sender, RoutedEventArgs e)
        {
            _view?.Refresh();
        }

        private void FilterToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            _view?.Refresh();
        }

        private bool FilterPredicate(Models.FormItem item)
        {
            if (item == null) return false;

            // Если ни одна кнопка не нажата – показываем всё (или ничего – решите сами)
            bool showErrors = buttonErrors.IsChecked == true;
            bool showWarnings = buttonWarnings.IsChecked == true;
            bool showMessages = buttonMessages.IsChecked == true;

            if (!showErrors && !showWarnings && !showMessages)
                return true;    // показывать все, если ничего не выбрано (можно изменить на false)

            // Проверяем соответствие кода
            return (showErrors && item.classError == 3) ||
                   (showWarnings && item.classError == 2) ||
                   (showMessages && item.classError == 1);
        }


        private void buttonCheck_Click(object sender, RoutedEventArgs e)
        {
            var pagesRules = new List<IRule> 
            { 
                new PlanTraccers(), // Проверяет трасеры
                new PageName(), // Проверяет имя старницы
                new Terminals(), // Проверяет клеммы
                new Cables(), // Проверяет кабели
               
            };
            var pagesDocRules = new List<IRule>
            {
                 new DocumentMasters(), // Проверка на дубликаты
            };

            Debug.WriteLine(_app.ActiveDocument);

            // Страницу
            if (comboDataSource.SelectedIndex == 0)
            {
                Page page = _app.ActivePage;

                Results.Validation validation = new Results.Validation(pagesRules,page);
                dataGrid.ItemsSource = validation.Validate();
                
            }
            else if (comboDataSource.SelectedIndex == 1) 
            {
                MessageBoxResult mresult = System.Windows.MessageBox.Show("Проверка всех страниц может занять много времени, вы точно хотите выполнить проверку всего документа?","Подтверждение",MessageBoxButton.YesNo,MessageBoxImage.Question);

                if (mresult != MessageBoxResult.Yes)
                    return;

                List<Models.FormItem> items = new List<Models.FormItem>();
                foreach (Page page in _app.ActiveDocument.Pages)
                {
                    Results.Validation validation = new Results.Validation(pagesRules, page);
                    items.AddRange(validation.Validate());
                }
                // Временная заглушка
                Results.Validation validation2 = new Results.Validation(pagesDocRules, _app.ActivePage);
                items.AddRange(validation2.Validate());
                dataGrid.ItemsSource = items;
            }
        }
  
        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Получаем объект, на котором произошёл двойной клик
            Models.FormItem selectedItem = dataGrid.SelectedItem as Models.FormItem; // замените на ваш класс
            if (selectedItem != null)
            {
                if (selectedItem.Shape == null || selectedItem.Page == null)
                    return;

                Navigate.NavigateToShape(_app, selectedItem.Shape);
                
            }
        }
    }


}
