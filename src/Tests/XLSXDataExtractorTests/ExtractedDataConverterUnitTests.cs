using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
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
        public void GenerateDataTableValidTest()
        {
            var dataTable = ExtractedDataConverter.GenerateDataTable(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());

            Assert.That(dataTable.Columns.Count, Is.EqualTo(10));
            Assert.That(dataTable.Rows.Count, Is.EqualTo(10));

            for (int i = 0; i < 10; i++)
            {
                DataColumn column = dataTable.Columns[i];
                Assert.That(column.ColumnName, Is.EqualTo("Test" + i));
            }
        }

        [Test]
        public void ConvertToWorksheetFromDataTableValidTest()
        {
            var dataTable = ExtractedDataConverter.GenerateDataTable(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());
            var worksheet = ExtractedDataConverter.ConvertToWorksheet(dataTable);

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
            IEnumerable<IEnumerable<KeyValuePair<string, object>>> nullObj = null;
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToWorksheet(nullObj));
        }

        [Test]
        public void ConvertToWorksheetNullDataTablenTest()
        {
            DataTable nullObj = null;
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToWorksheet(nullObj));
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
            string expectedCSV = File.ReadAllText(Path.Combine(executingAssemblyPath,"Files", "ExpectedCSV-CommaEscaped.txt"));

            string actualCSV = ExtractedDataConverter.ConvertToCSV(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());

            Assert.That(actualCSV, Is.EqualTo(expectedCSV));
        }

        [Test]
        public void ConvertToCSVFromDataTable()
        {
            string expectedCSV = File.ReadAllText(Path.Combine(executingAssemblyPath, "Files", "ExpectedCSV-CommaEscaped.txt"));

            var dataTable = ExtractedDataConverter.GenerateDataTable(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());         
            string actualCSV = ExtractedDataConverter.ConvertToCSV(dataTable);

            Assert.That(actualCSV, Is.EqualTo(expectedCSV));
        }

        [Test]
        public void ConvertToCSVFromDataTable_TabDelimited_NotEscaped()
        {
            string expectedCSV = ExpectedStrings.ExpectedCSV_TabDelimited_NotEscaped;

            var dataTable = ExtractedDataConverter.GenerateDataTable(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedData());
            string actualCSV = ExtractedDataConverter.ConvertToDelimitedString(dataTable, "\t", false);

            Assert.That(actualCSV, Is.EqualTo(expectedCSV));
        }

        [Test]
        public void ConvertToCSVNullCollectionTest()
        {
            IEnumerable<IEnumerable<KeyValuePair<string, object>>> nullObj = null;
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToCSV(nullObj));
        }

        [Test]
        public void ConvertToCSVNullDataTableTest()
        {
            DataTable nullObj = null;
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToCSV(nullObj));
        }

        [Test]
        public void ConvertToCSVCollectionWithNullFieldNames()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => ExtractedDataConverter.ConvertToCSV(XLSXDataExtractorTestHelper.GenTwoDimensionalCollectionOfExtractedDataWithNullFieldNames()));
            Assert.That(ex.Message, Is.EqualTo("An ExtractedData object at position 0,5 has a null field name, null field names cannot be added.\r\nParameter name: extractedData"));
        }
    }
}
