using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converters
{
    public class DefaultOffice : IOffice
    {
        public void ConvertToPdf(string filePath)
        {
            throw new NotInstalledOfficeException();
        }
    }
}
