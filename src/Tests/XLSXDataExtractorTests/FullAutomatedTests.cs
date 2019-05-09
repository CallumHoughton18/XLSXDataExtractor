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
        public void ValidWorksheetGenerationIntegrationTest()
        {
            DataExtractor dataExtractor = new DataExtractor(Path.Combine(executingAssemblyPath, "Files", "IntegrationTestExample.xlsx"));

            var extractionRequests = new List<ExtractionRequest>() { new ExtractionRequest("SalesRep", 2, 3), new ExtractionRequest("SalesRepID", 2, 2) };
            var extracted = dataExtractor.RetrieveDataCollectionFromAllWorksheets<object>(extractionRequests);

            var generatedWorksheet = ExtractedDataConverter.ConvertToWorksheet(extracted);

            Assert.That(generatedWorksheet.Cell(1, 1).Value.ToString(), Is.EqualTo("SalesRep"));
            Assert.That(generatedWorksheet.Cell(2, 1).Value.ToString(), Is.EqualTo("Jane"));
            Assert.That(generatedWorksheet.Cell(3, 1).Value.ToString(), Is.EqualTo("Ashish"));
            Assert.That(generatedWorksheet.Cell(4, 1).Value.ToString(), Is.EqualTo("John"));

            Assert.That(generatedWorksheet.Cell(1, 2).Value.ToString(), Is.EqualTo("SalesRepID"));
            Assert.That(generatedWorksheet.Cell(2, 2).Value.ToString(), Is.EqualTo("456"));
            Assert.That(generatedWorksheet.Cell(3, 2).Value.ToString(), Is.EqualTo("789"));
            Assert.That(generatedWorksheet.Cell(4, 2).Value.ToString(), Is.EqualTo("123"));
        }

        [Test]
        public void ValidCSVGenerationIntegrationTest()
        {
            DataExtractor dataExtractor = new DataExtractor(Path.Combine(executingAssemblyPath, "Files", "IntegrationTestExample.xlsx"));

            var extractionRequests = new List<ExtractionRequest>() { new ExtractionRequest("SalesRep", 2, 3), new ExtractionRequest("SalesRepID", 2, 2) };
            var extracted = dataExtractor.RetrieveDataCollectionFromAllWorksheets<object>(extractionRequests);

            var generatedCSVText = ExtractedDataConverter.ConvertToCSV(extracted);
            Assert.That(generatedCSVText, Is.EqualTo(File.ReadAllText(Path.Combine(executingAssemblyPath, "Files", "IntegrationTestExpectedCsv.txt"))));
        }
    }
}
