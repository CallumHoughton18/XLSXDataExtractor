using ClosedXML.Excel;
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
    public class ExtractedDataConverterUnitTests
    {
        string executingAssemblyPath;

        [OneTimeSetUp]
        public void InitialOneTimeSetUp()
        {
            executingAssemblyPath = AppDomain.CurrentDomain.BaseDirectory;
        }

        [Test]
        public void ConvertToWorksheetValidTest()
        {
            var worksheet = ExtractedDataConverter.ConvertToWorksheet(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());

            var cellsUsed = worksheet.CellsUsed();

            bool emptyCellsFromDataTable = cellsUsed.Any(x => x.Address.ColumnNumber == 6 && x.Address.RowNumber > 1);
            Assert.That(emptyCellsFromDataTable, Is.False);
            foreach (var cell in cellsUsed)
            {
                //generated table should be 12 cells by 12 cells start from cell 0,0.
                Assert.That(cell.Address.RowNumber, Is.LessThan(12));
                Assert.That(cell.Address.ColumnNumber, Is.LessThan(12));
            }
        }

        [Test]
        public void ConvertToWorksheetNullCollectionTest()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToWorksheet(null));
        }

        [Test]
        public void ConvertToWorksheetCollectionWithNullFieldNames()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToWorksheet(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedDataWithNullFieldNames()));
            Assert.That(ex.Message, Is.EqualTo("An ExtractedData object at position 0,5 has a null field name, null field names cannot be added.\r\nParameter name: extractedData"));
        }

        [Test]
        public void ConvertToCSVValid()
        {
            string expectedCSV = File.ReadAllText(Path.Combine(executingAssemblyPath,"Files", "ExpectedCSV.txt"));

            string actualCSV = ExtractedDataConverter.ConvertToCSV(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());

            Assert.That(actualCSV, Is.EqualTo(expectedCSV));
        }

        [Test]
        public void ConvertToCSVNullCollectionTest()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToCSV(null));
        }

        [Test]
        public void ConvertToCSVCollectionWithNullFieldNames()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToCSV(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedDataWithNullFieldNames()));
            Assert.That(ex.Message, Is.EqualTo("An ExtractedData object at position 0,5 has a null field name, null field names cannot be added.\r\nParameter name: extractedData"));
        }
    }
}
