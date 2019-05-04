using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXDataExtractor.Models
{
    public abstract class ExtractedDataBase
    {
        public abstract Type Type { get; }
    }
}
