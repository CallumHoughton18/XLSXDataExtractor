using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXDataExtractor;

namespace XLSXDataExtractorTests
{
    public class ExtractedDataConverterTests
    {
        string executingAssemblyPath;

        [OneTimeSetUp]
        public void InitialOneTimeSetUp()
        {
            executingAssemblyPath = AppDomain.CurrentDomain.BaseDirectory;
        }

        [Test]
        public void ConvertToXLSXValidTest()
        {
            var sut = new ExtractedDataConverter();
            sut.ConvertToNewXLSX(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData(), Path.Combine(executingAssemblyPath,"TestGen.XLSX"));
        }

        [Test]
        public void ConvertToCSVValid()
        {
            var sut = new ExtractedDataConverter();
            sut.ConvertToCSV(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData(), Path.Combine(executingAssemblyPath, "TestGen.CSV"));
        }
    }
}
