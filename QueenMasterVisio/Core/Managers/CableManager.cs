using Microsoft.Office.Interop.Visio;
using QueenMasterVisio.Core.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace QueenMasterVisio.Core.Managers
{
    internal class CableManager
    {
        /// <summary>
        /// Создание и манипуляции с кабелем
        /// </summary>
        Master master;
        Page page;
        public bool ready = false;
        public CableManager(Page page) 
        {
            try
            {
                this.master = page.Document.Masters.get_ItemU("Cable");
                this.ready = true;
                this.page = page;
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось получить мастер кабеля");
                throw;
            } 
        }
        public void Create(string type, string way, string toDevice, double x, double y)
        {
            Shape shape = this.page.Drop(master, x, y);
            shape.SetFormula("Prop.type", type);
            shape.SetFormula("Prop.Name", way);
            shape.SetFormula("Prop.Device", toDevice);
        }
    }
}
