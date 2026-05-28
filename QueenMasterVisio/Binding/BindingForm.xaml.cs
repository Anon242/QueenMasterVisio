using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Page = Microsoft.Office.Interop.Visio.Page;
using Shape = Microsoft.Office.Interop.Visio.Shape;

namespace QueenMasterVisio.Binding
{
    /// <summary>
    /// Логика взаимодействия для BindingForm.xaml
    /// </summary>
    public partial class BindingForm : UserControl
    {
        private Document doc;
        private Page thisPage;

        public BindingForm(Page page)
        {
            InitializeComponent();
            this.doc = page.Document;
            this.thisPage = page;
            GetData();
        }

        private void RebuildPageUserCells()
        {
            if (RightListBox.Items.Count > 0) 
            {
                string result = string.Join(";", RightListBox.Items.Cast<string>());
                thisPage.SetUserCell("deviceList", result);
                string firstElem = RightListBox.Items[0].ToString();
                if (firstElem[0] == 'G')
                    thisPage.SetUserCell("pageCode", "Device");
                else if (firstElem[0] == 'L')
                    thisPage.SetUserCell("pageCode", "Light");

            }
        }

        private void AddToRight()
        {
            string device = LeftTreeView.SelectedItem as string;
            if (device != null)
            {
                if (!RightListBox.Items.Cast<string>().Contains(device))
                {
                    // Можно добавить только того же типа
                    if (RightListBox.Items.Count > 0)
                    {
                        string firstElem = RightListBox.Items[0].ToString();
                        if (firstElem[0] == device[0])
                        {
                            RightListBox.Items.Add(device);

                        }
                    }
                    else
                    {
                        RightListBox.Items.Add(device);

                    }
                    RebuildPageUserCells();
                }
            }
        }


        private void DeleteRightElem()
        {
            if (RightListBox.SelectedItem != null)
            {
                RightListBox.Items.Remove(RightListBox.SelectedItem);
                RebuildPageUserCells();
            }
        }

        private void CheckDevices(string device)
        {
            bool isFind = false;
            foreach (TreeViewItem item in LeftTreeView.Items)
            {
                foreach (var viewItem in item.Items)
                {
                    if (device == viewItem)
                        isFind = true;
                    break;
                }
                if (isFind)
                    break;
            }
            if (!isFind)
            {

            }
        }

        private void GetData()
        {
            if (this.doc == null)
                return;

            foreach (Page page in doc.Pages)
            {
                if(page.IsPlanPage())
                {
                    List<Shape> devices = GetDeviceAndLightObjects(page);
                    AddElementsToTreeView(page.Name, devices);
                }
            }
            if (!thisPage.HasCell("User.pageCode"))
                thisPage.SetUserCell("pageCode", "");
            if (!thisPage.HasCell("User.deviceList"))
                thisPage.SetUserCell("deviceList", "");

            string [] result = thisPage.GetCellFormulaU("User.deviceList").Split(';');
            if(result.Length > 0)
                foreach (string device in result)
                    if(!string.IsNullOrEmpty(device))
                        RightListBox.Items.Add(device);
            
        }

        private void AddElementsToTreeView(string name, List<Shape> shapes)
        {
            var newChild = new TreeViewItem { Header = name };
            shapes = shapes.OrderBy(s => s.NameU.Contains("Device") ? "G" : "L" + s.GetCellResultString("Prop.number")).ToList();
            foreach (Shape shape in shapes) 
            {
                string number = shape.GetCellResultString("Prop.number");
                string deviceName = "";
                if (shape.NameU.Contains("Device"))
                    deviceName = "G" + number;
                else
                    deviceName = "L" + number;

                newChild.Items.Add(deviceName + " - " + shape.NameU);
            }
            LeftTreeView.Items.Add(newChild);
        }

        private List<Shape> GetDeviceAndLightObjects(Page page)
        {
            List <Shape> result = new List<Shape>();
            foreach (Shape shape in page.Shapes)
            {
                if(shape.NameU.Contains("Device") || shape.NameU.Contains("Light"))
                {
                    result.Add(shape);
                }
            }
            return result;
        }

        private void MoveRightButton_Click(object sender, RoutedEventArgs e)
        {
            AddToRight();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteRightElem();
        }

        private void LeftDoubleClick(object sender, MouseButtonEventArgs e)
        {
            AddToRight();
        }

        private void RightClick(object sender, MouseButtonEventArgs e)
        {
            DeleteRightElem();
        }
    }
}
