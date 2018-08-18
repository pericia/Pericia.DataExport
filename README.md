# Pericia.DataExport

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

### Using Attributes on your model

Add the attribute `ExportColumn` on the properties you want to export. The attributes contains 2 properties : `Title` (the name of the column) ans `Order` (tu sort your columns)

    public class SampleData
    {
        [ExportColumn(Title = "Number", Order = 1)]
        public int IntData { get; set; }

        [ExportColumn(Title = "Text", Order = 2)]
        public string TextData { get; set; }
    }

### Create your exporter

You can use either `CsvDataExporter` or `XlsxDataExporter` :

	var data = new List<SampleData>()
	{
		new SampleData{ IntData=5, TextData="Hello"},
		new SampleData{ IntData=20, TextData="Yoo"},
		new SampleData{ IntData=10, TextData="This is some text"},
	};

	var csvExporter = new CsvDataExporter();
	var csvResult = exporter.Export(data);
	
	var xlsxExporter = new XlsxDataExporter();
	var xlsxResult = exporter.Export(data);

### Create Xlsx file with several sheets

While the csv exporter will only allow you to export one set of data, with the xlsx exporter you can create several sheets with different data on each.

	var xlsxExporter = new XlsxDataExporter();
	xlsxExporter.AddSheet(data1, name="sheet title 1");
	xlsxExporter.AddSheet(data2, name="sheet title 2");
	var xlsxResult = exporter.GetFile();

### Result

The exporters will output a `MemoryStream`. You can directly save it to a file, or return it in a `FileResult` in an MVC website.
