using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXDataExtractor;
using XLSXDataExtractor.Models;

namespace XLSXDataExtractorTests
{
    class FullAutomatedTests
    {
        string executingAssemblyPath;

        [OneTimeSetUp]
        public void InitialOneTimeSetUp()
        {
            executingAssemblyPath = AppDomain.CurrentDomain.BaseDirectory;
        }

        [Test]
        public void ValidIntegrationTest()
        {
            DataExtractor dataExtractor = new DataExtractor(Path.Combine(executingAssemblyPath, "Files", "IntegrationTestExample.xlsx"));

            var extractionRequests = new List<ExtractionRequest>() { new ExtractionRequest("SalesRep", 2, 3), new ExtractionRequest("SalesRepID", 3, 2) };

            var extracted = dataExtractor.RetrieveDataFromAllWorksheetsInWorkbook<object>(extractionRequests);

            XLWorkbook workbook = new XLWorkbook();
            workbook.AddWorksheet(ExtractedDataConverter.ConvertToWorksheet(extracted));
            workbook.SaveAs(Path.Combine(executingAssemblyPath, "IntegrationTest.xlsx"));
        }
    }
}
