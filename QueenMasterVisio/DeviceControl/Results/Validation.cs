using QueenMasterVisio.DeviceControl.Rules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Page = Microsoft.Office.Interop.Visio.Page;
using Visio = Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace QueenMasterVisio.DeviceControl.Results
{
    internal class Validation
    {
        List<IRule> rules = new List<IRule>();
        Page page { get; set; }
        ValidationEngine engine;
        public Validation(List<IRule> _rules, Page _page)
        {
            rules = _rules;
            page = _page;
            engine = new ValidationEngine(page);
        }

        public List<Models.FormItem> Validate()
        {
            try
            {
                List<Models.FormItem> items = new List<Models.FormItem>();
                foreach (var item in rules)
                {
                    foreach (ValidationResult result in item.Validate(page, engine))
                    {
                        items.Add(new Models.FormItem
                        {
                            Code = (int)result.Severity,
                            Description = result.Message,
                            Shape = result.TargetShape,
                            Page = result.TargetPage,

                        });
                    }
                }
                return items;
            }
            catch (Exception)
            {

                return new List<Models.FormItem>();
            }
        }
    }
}
