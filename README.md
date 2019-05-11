# XLSXDataExtractor

An extension library for ClosedXML to allow for quick and easy bulk data extraction from Excel Workbooks into various formats.

[![Build Status](https://travis-ci.org/CallumHoughton18/XLSXDataExtractor.svg?branch=master)](https://travis-ci.org/CallumHoughton18/XLSXDataExtractor)

**NuGets**

| Name          			   | Link          														   |
| -------------------------|:---------------------------------------------------------------------:|
| XLSX_Data_Extractor      | [NuGet Link](https://www.nuget.org/packages/XLSX_Data_Extractor/1.0.0)|
| XLSXDataExtractor.Common | [NuGet Link](https://www.nuget.org/packages/XLSXDataExtractor.Common/)|        

## What's Included?

Two components are included. The DataExtractor class which is used to extract data out from Worksheets of a loaded Workbook, and the ExtractedDataConverter static class which is used to convert the extracted data into various formats, such as a CSV or XLSX document.

### DataExtractor Usage

The DataExtractor contains multiple overloads for single data extraction using the worksheets known name, column, and row of the data. Or, an ExtractionRequest object can be passed to it containing the same information. For bulk data extraction a collection of ExtractionRequest objects can be passed to it, an example of this is shown below.

```C#
DataExtractor dataExtractor = new DataExtractor(Path.Combine(executingAssemblyPath, "Files", "IntegrationTestExample.xlsx"));

var extractionRequests = new List<ExtractionRequest>() { new ExtractionRequest("SalesRep", 2, 3), new ExtractionRequest("SalesRepID", 2, 2) };

var extracted = dataExtractor.RetrieveDataCollectionFromAllWorksheets<object>(extractionRequests);
```

### ExtractedDataConverter Usage

This is a static class than converts a two dimensional collection of the KeyValuePair objects produced from the various extraction methods into a string of CSV text or a IXLWorksheet to be used with the ClosedXML library. 