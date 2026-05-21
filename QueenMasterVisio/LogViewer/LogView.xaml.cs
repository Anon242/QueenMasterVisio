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
using System.IO;

namespace QueenMasterVisio.LogViewer
{
    /// <summary>
    /// Логика взаимодействия для LogView.xaml
    /// </summary>
    public partial class LogView : UserControl
    {
        public LogView(List<string> filesPath)
        {
            InitializeComponent();
            if(filesPath.Count > 0)
            {
                List<Models.FormItem> items = new List<Models.FormItem>();
                foreach (string file in filesPath.Reverse<string>())
                {
                    if (File.Exists(file))
                    {
                        var item = new Models.FormItem();
                        string text = File.ReadAllText(file);
                        item.name = text.Split('\n')[0];
                        item.comment = text.Split('\n')[1];
                        item.author = System.IO.Path.GetFileNameWithoutExtension(file).Split('_')[1];
                        item.time = new DirectoryInfo(file).LastWriteTime.ToString("yyyy-MM-dd HH:mm");

                        items.Add(item);
                    }
                    
                }
                var sortedItems = items.OrderByDescending(i => i.time).ToList();
                LogDataGrid.ItemsSource = sortedItems;
                
            }
        }
    }
}
