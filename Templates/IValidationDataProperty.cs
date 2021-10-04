using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapping
{
    public interface IValidationDataProperty
    {
        /// <summary>
        /// return true if the value is valid, return false if the value is not valid.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        bool CheckPropertyValue(string value);
    }
}
