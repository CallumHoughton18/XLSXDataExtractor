﻿using NUnit.Framework;
using System;
using System.IO;
using XLSXDataExtractor;

namespace XLSXDataExtractorTests
{
    public class DataExtractorUnitTests
    {
        string executingAssemblyPath;

        [OneTimeSetUp]
        public void InitialOneTimeSetUp()
        {
            executingAssemblyPath = AppDomain.CurrentDomain.BaseDirectory;
        }

        [Test]
        public void InitializationValidWorkbookTest()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);
        }

        [Test]
        public void InitializationNoWorkbookFound()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "filedoesnotexist.xlsx");

            DataExtractor sut;
            var ex = Assert.Throws<FileNotFoundException>(() => sut = new DataExtractor(workbookPath));
        }

        [Test]
        public void InitializationInvalidFile()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "expectedCSV.txt");

            DataExtractor sut;
            var ex = Assert.Throws<ArgumentException>(() => sut = new DataExtractor(workbookPath));
            StringAssert.Contains("Extension 'txt' is not supported. Supported extensions are '.xlsx', '.xslm', '.xltx' and '.xltm'.", ex.Message);
        }

        [Test]
        public void RetrieveDataIntOverloadValidWorksheet()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var extractedData = sut.RetrieveDataFromWorkbook<string>(1, "TestField", 1, 1);

            Assert.That(extractedData.Key, Is.EqualTo("TestField"));
            Assert.That(extractedData.Value, Is.EqualTo("TestValue1"));
        }

        [Test]
        public void RetrieveDataIntOverloadInValidWorksheet()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => sut.RetrieveDataFromWorkbook<string>(33, "TestField", 1, 1));
            Assert.That(ex.Message, Is.EqualTo("Not within range of worksheets. Worksheets count: 3\r\nParameter name: workSheetNum"));
        }

        [Test]
        public void RetrieveDataFromWorksheetNameInvalidName()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var ex = Assert.Throws<NullReferenceException>(() => sut.RetrieveDataFromWorkbook<string>("notavalidworksheetname", "TestField", 1, 1));
            Assert.That(ex.Message, Is.EqualTo($"Worksheet notavalidworksheetname does not exist in {Path.GetFileName(workbookPath)}"));
        }

        [Test]
        public void RetrieveDataValidAsString()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var extractedData = sut.RetrieveDataFromWorkbook<string>("Sheet1", "TestField", 1, 1);

            Assert.That(extractedData.Key, Is.EqualTo("TestField"));
            Assert.That(extractedData.Value, Is.EqualTo("TestValue1"));
        }

        [Test]
        public void RetrieveDataValidAsDouble()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var extractedData = sut.RetrieveDataFromWorkbook<double>("Sheet3", "TestField", 1, 2);

            Assert.That(extractedData.Key, Is.EqualTo("TestField"));
            Assert.That(extractedData.Value, Is.EqualTo(3.33D));
        }

        [Test]
        public void RetrieveDataInvalidConversion()
        {
            //attempt to retrieve data from cell which cannot be cast as a double, only as a string.

            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var ex = Assert.Throws<NullReferenceException>(() => sut.RetrieveDataFromWorkbook<double>("Sheet1", "TestField", 1, 1));
            Assert.That(ex.Message, Is.EqualTo("No value in row:1 col:1 in Sheet1 of type Double"));
        }

        [Test]
        public void RetrieveDataAsStringNoValueInCell()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var extractedData = sut.RetrieveDataFromWorkbook<string>("Sheet1", "TestField", 1, 5);
            Assert.That(extractedData.Key, Is.EqualTo("TestField"));
            Assert.That(extractedData.Value, Is.EqualTo(""));
        }

        [Test]
        public void RetrieveDataAsDoubleNoValueInCell()
        {
            var workbookPath = Path.Combine(executingAssemblyPath, "Files", "testWorkbook.xlsx");
            var sut = new DataExtractor(workbookPath);

            var ex = Assert.Throws<NullReferenceException>(() => sut.RetrieveDataFromWorkbook<double>("Sheet1", "TestField", 1, 10));
            Assert.That(ex.Message, Is.EqualTo("No value in row:1 col:10 in Sheet1 of type Double"));
        }
    }
}
