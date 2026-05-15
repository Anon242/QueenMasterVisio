using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace QueenMasterVisio.DeviceControl
{
    /// <summary>
    /// Логика взаимодействия для DeviceControlForm.xaml
    /// </summary>
    public partial class DeviceControlForm : System.Windows.Controls.UserControl
    {
        private ICollectionView _view;
        public DeviceControlForm()
        {
            InitializeComponent();
            LoadData();
            SetupFiltering();
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
        private void LoadData()
        {
            // Пример данных: список объектов
            var items = new List<Models.FormItem>()
            {
                new Models.FormItem { Code = 101, Description = "Красный квадрат", Shape = "Квадрат", Page = 12 },
                new Models.FormItem { Code = 102, Description = "Синий кругкругкругкруг кругкруг", Shape = "Круг", Page = 15 },
                new Models.FormItem { Code = 103, Description = "Зелёный треугольник", Shape = "Треугольник", Page = 20 },
                new Models.FormItem { Code = 201, Description = "Красный квадрат", Shape = "Квадрат", Page = 12 },
                new Models.FormItem { Code = 202, Description = "Синий кругкругкругкруг кругкруг", Shape = "Круг", Page = 15 },
                new Models.FormItem { Code = 203, Description = "Зелёный треугольник", Shape = "Треугольник", Page = 20 },
                new Models.FormItem { Code = 201, Description = "Красный квадрат", Shape = "Квадрат", Page = 12 },
                new Models.FormItem { Code = 102, Description = "Синий кругкругкругкруг кругкруг", Shape = "Круг", Page = 15 },
                new Models.FormItem { Code = 303, Description = "Зелёный треугольник", Shape = "Треугольник", Page = 20 },
                new Models.FormItem { Code = 301, Description = "Красный квадрат", Shape = "Квадрат", Page = 12 },
                new Models.FormItem { Code = 302, Description = "Синий кругкругкругкруг кругкруг", Shape = "Круг", Page = 15 },
                new Models.FormItem { Code = 303, Description = "Зелёный треугольник", Shape = "Треугольник", Page = 20 }
            };

            dataGrid.ItemsSource = items;
        }

        private void buttonCheck_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine(comboDataSource.SelectedIndex);
            // Страницу
            if(comboDataSource.SelectedIndex == 0)
            {
                // Делаем запрос на проверку страницы
            }
        }
    }


}
