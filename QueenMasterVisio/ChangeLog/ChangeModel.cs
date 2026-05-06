using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueenMasterVisio.ChangeLog
{

    internal class ChangeModel
    {
        public string name;
        public string description;
        public DateTime date;
        public string log;
        public string version;

        public ChangeModel(string name, string description, DateTime date, string log, string version)
        {
            this.name = name;
            this.description = description;
            this.date = date;
            this.log = log;
            this.version = version;
        }
    }
}
