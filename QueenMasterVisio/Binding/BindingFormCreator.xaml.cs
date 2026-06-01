using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using QueenMasterVisio.Core.Managers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Visio.Page;
using Shape = Microsoft.Office.Interop.Visio.Shape;

namespace QueenMasterVisio.Binding
{
    /// <summary>
    /// Логика взаимодействия для BindingForm.xaml
    /// </summary>
    public partial class BindingFormCreator : UserControl
    {
        private Document doc;
        private Page thisPage;
        private System.Windows.Window window;

        public BindingFormCreator(Page page, System.Windows.Window window)
        {
            InitializeComponent();
            this.doc = page.Document;
            this.thisPage = page;
            this.window = window;
            GetData();
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
                }
               
            }
            CheckCreateButton();
        }

        private void CheckCreateButton()
        {
            if (RightListBox.Items.Count > 0 && !string.IsNullOrEmpty(PageNameTextBox.Text)) 
                CreateButton.IsEnabled = true;
            else
                CreateButton.IsEnabled = false;
        }

        private void DeleteRightElem()
        {
            if (RightListBox.SelectedItem != null)
            {
                RightListBox.Items.Remove(RightListBox.SelectedItem);
            }
            CheckCreateButton();
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

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (RightListBox.Items.Count > 0 && !string.IsNullOrEmpty(PageNameTextBox.Text))
            {
                string result = string.Join(";", RightListBox.Items.Cast<string>());

                // Создаем новую страницу
                List<string> allDevices = new List<string>();
                List<string> allObjects = new List<string>();
                foreach (string item in RightListBox.Items.Cast<string>())
                {
                    allDevices.Add(item.Split('-')[0].Trim());
                    allObjects.Add(item.Split('-')[1].Trim());
                }

                Page newPage = DocumentManager.CreateNewPage(string.Join(", ", allDevices) + " " + PageNameTextBox.Text);

                if (newPage.Name[0] == 'G')
                {
                    // Делаем что она была перед первым светом
                    foreach (Page _page in doc.Pages)
                    {
                        Regex regexLight = new Regex(@"^L\d");
                        if (regexLight.IsMatch(_page.Name))
                        {
                            newPage.Index = _page.Index;
                            break;
                        }
                    }
                }

                if (!newPage.HasCell("User.pageCode"))
                    newPage.SetUserCell("pageCode", "");
                if (!newPage.HasCell("User.deviceList"))
                    newPage.SetUserCell("deviceList", "");

                newPage.SetUserCell("deviceList", result);
                string firstElem = RightListBox.Items[0].ToString();
                if (firstElem[0] == 'G')
                    newPage.SetUserCell("pageCode", "Device");
                else if (firstElem[0] == 'L')
                    newPage.SetUserCell("pageCode", "Light");

                foreach (string item in allObjects)
                {
                    try
                    {
                        // А еще надо проверить на существование линков
                        string pageName = GetPageNameForShape(item);
                        
                        Shape shape = doc.Pages[pageName].Shapes.ItemU[item];
                        if (shape != null)
                        {
                            //Microsoft.Office.Interop.Visio.Hyperlink hlink = shape.Hyperlinks.Add();
                            //hlink.SubAddress = newPage.NameU;
                            Microsoft.Office.Interop.Visio.Hyperlink hyperlink = shape.AddHyperlink();
                            hyperlink.SubAddress = newPage.NameU; 
                            hyperlink.Description = newPage.Name;
                        }
                    }
                    catch
                    {

                    }
                }

                // Делаем так чтобы девайс встал нормально

                
                  
               

                 

                window.Close();
            }
        }
        private string GetPageNameForShape(string shapeName)
        {
            foreach (TreeViewItem root in LeftTreeView.Items)
            {
                foreach (object child in root.Items)
                {
                    string childText = child.ToString(); 
                                                         
                    if (childText.Contains(shapeName))
                    {
                        return root.Header.ToString(); 
                    }
                }
            }
            return null;
        }

        private void PageNameTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            CheckCreateButton();
        }
    }
}
