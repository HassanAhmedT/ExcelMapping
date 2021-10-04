using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapping
{
    public interface IParsingDataProperty
    {
        /// <summary>
        /// casting the string value to the desired data type
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        object ParsePropertyValue(string value);
    }
}
