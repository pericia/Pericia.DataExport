# Pericia.DataExport

[![Build status](https://dev.azure.com/glacasa/GithubBuilds/_apis/build/status/Pericia.DataExport-CI)](https://dev.azure.com/glacasa/GithubBuilds/_build/latest?definitionId=62)

Pericia.DataExport is a dotnet library to export any `IEnumerable<T>` to an xlsx or csv file.

## Install 

The library `Pericia.DataExport` is available on Nuget : [![NuGet](https://img.shields.io/nuget/v/Pericia.DataExport.svg)](https://www.nuget.org/packages/Pericia.DataExport/)

You can install it with the following command line in your Package Manager Console :

	Install-Package Pericia.DataExport

Or with dotnet core :

	 dotnet add package Pericia.DataExport 

## How to use

The data exported will be a file with column titles on the first line, and all your data on the following lines.

To export your data, you will need to define how it will be exported

### Export using Attributes on your model

Add the attribute `ExportColumn` on the properties you want to export. The attributes contains 2 properties : `Title` (the name of the column) ans `Order` (tu sort your columns)
```cs
    public class SampleData
    {
        [ExportColumn(Title = "Number", Order = 1)]
        public int IntData { get; set; }

        [ExportColumn(Title = "Text", Order = 2)]
        public string TextData { get; set; }
    }
```
### Create your exporter

You can use either `CsvDataExporter` or `XlsxDataExporter` :
```cs
	var data = new List<SampleData>()
	{
		new SampleData{ IntData=5, TextData="Hello"},
		new SampleData{ IntData=20, TextData="Yoo"},
		new SampleData{ IntData=10, TextData="This is some text"},
	};

	var csvExporter = new CsvDataExporter();
	var csvResult = csvExporter.Export(data);
	
	var xlsxExporter = new XlsxDataExporter();
	var xlsxResult = xlsxExporter.Export(data);
```
### Create Xlsx file with several sheets

While the csv exporter will only allow you to export one set of data, with the xlsx exporter you can create several sheets with different data on each.
```cs
	var xlsxExporter = new XlsxDataExporter();
	xlsxExporter.AddSheet(data1, name="sheet title 1");
	xlsxExporter.AddSheet(data2, name="sheet title 2");
	var xlsxResult = xlsxExporter.GetFile();
```
### Export using SQL Data Reader

If you want to export a query result without binding it to a model, you can use an `SqlDataReader` :
```cs
	using (SqlConnection connection = new SqlConnection(connectionString))
	{
		SqlCommand command = new SqlCommand("SELECT OrderID, CustomerID FROM dbo.Orders", connection);
		connection.Open();
		SqlDataReader reader = command.ExecuteReader();
		var xlsxExporter = new XlsxDataExporter();
		var xlsxResult = xlsxExporter.Export(reader);
	}
```
### Result

The exporters will output a `MemoryStream`. You can directly save it to a file, or return it in a `FileResult` in an MVC website.
