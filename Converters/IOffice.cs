﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converters
{
    public interface IOffice
    {
        void ConvertToPdf(string filePath);
    }
}
