﻿using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace Pericia.DataExport
{
    public class ExportTest
    {
        [Fact]
        public void CsvExportTest()
        {
            var exporter = new CsvDataExporter();

            var data = new List<SampleData>()
            {
                new SampleData( 5, "Hello", true),
                new SampleData(20,"A,B;C", false),
                new SampleData(10, "A\"B,C", true),
            };

            var exportResult = exporter.Export(data);

            var reader = new StreamReader(exportResult);
            Assert.Equal(@"Number;Text;Bool", reader.ReadLine());
            Assert.Equal(@"5;Hello;True", reader.ReadLine());
            Assert.Equal(@"20;""A,B;C"";False", reader.ReadLine());
            Assert.Equal(@"10;""A""""B,C"";True", reader.ReadLine());

        }

        [Fact]
        public void ConverterTest()
        {
            var exporter = new CsvDataExporter();

            exporter.AddDataConverter<bool>(b => b ? "👍" : "❌");

            var data = new List<SampleData>()
            {
                new SampleData( 5, "Hello", true),
                new SampleData(20,"A,B;C", false),
                new SampleData(10, "A\"B,C", true),
            };

            var exportResult = exporter.Export(data);

            var reader = new StreamReader(exportResult);
            Assert.Equal(@"Number;Text;Bool", reader.ReadLine());
            Assert.Equal(@"5;Hello;👍", reader.ReadLine());
            Assert.Equal(@"20;""A,B;C"";❌", reader.ReadLine());
            Assert.Equal(@"10;""A""""B,C"";👍", reader.ReadLine());
        }
    }

    public class SampleData
    {
        public SampleData(int intData, string textData, bool boolData)
        {
            IntData = intData;
            TextData = textData;
            BoolData = boolData;
        }

        [ExportColumn(Title = "Number", Order = 1)]
        public int IntData { get; set; }

        [ExportColumn(Title = "Text", Order = 2)]
        public string TextData { get; set; }

        [ExportColumn(Title = "Bool", Order = 3)]
        public bool BoolData { get; set; }

    }
}
