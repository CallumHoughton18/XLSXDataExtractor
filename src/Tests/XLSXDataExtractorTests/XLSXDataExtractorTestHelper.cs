using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXDataExtractor.Models;

namespace XLSXDataExtractorTests
{
    public static class XLSXDataExtractorTestHelper
    {
        public static IEnumerable<IEnumerable<ExtractedData<object>>> GenTwoDimensionalCollectionOfExtractedData()
        {
            for (int i = 0; i < 10; i++)
            {
                List<ExtractedData<object>> col = new List<ExtractedData<object>>();

                for (int j = 0; j < 10; j++)
                {
                    if (j != 5)
                        col.Add(new ExtractedData<object>("Test" + j, j));

                    else
                        col.Add(new ExtractedData<object>("Test" + j, null));
                }

                yield return col;
            }
        }
    }
}
