using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapping
{
    internal class PropertyColumn
    {
        public PropertyColumn()
        {
            Validators = new List<IValidationDataProperty>();
            Parser = null;
        }

        public int ColumnNumber { get; set; }
        public string ColumnName { get; set; }
        public string PropertyName { get; set; }
        public IParsingDataProperty Parser { get; set; }
        public List<IValidationDataProperty> Validators { get; set; }
    }
}
