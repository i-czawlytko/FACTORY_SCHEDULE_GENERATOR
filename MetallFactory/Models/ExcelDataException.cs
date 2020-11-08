using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class ExcelDataException : Exception
    {
        public ExcelDataException(string message)
            : base(message)
                { }
    }
}
