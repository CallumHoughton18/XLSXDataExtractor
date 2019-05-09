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
        public static IEnumerable<IEnumerable<KeyValuePair<string, object>>> GenTwoDimensionalCollectionOfExtractedData()
        {
            for (int i = 0; i < 10; i++)
            {
                var col = new List<KeyValuePair<string, object>>();

                for (int j = 0; j < 10; j++)
                {
                    if (j != 5)
                        col.Add(new KeyValuePair<string, object>("Test" + j, j));

                    else
                        col.Add(new KeyValuePair<string, object>("Test" + j, null));
                }

                yield return col;
            }
        }

        public static IEnumerable<IEnumerable<KeyValuePair<string, object>>> GenTwoDimensionalCollectionOfExtractedDataWithNullFieldNames()
        {
            for (int i = 0; i < 10; i++)
            {
                var col = new List<KeyValuePair<string, object>>();

                for (int j = 0; j < 10; j++)
                {
                    if (j != 5)
                        col.Add(new KeyValuePair<string, object>("Test" + j, j));

                    else
                        col.Add(new KeyValuePair<string, object>(null, null));
                }

                yield return col;
            }
        }
    }
}
