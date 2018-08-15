using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace Pericia.DataExport.Tests
{
    public class ExportTest
    {
        [Fact]
        public void Test1()
        {
            var exporter = new DataExporter(ExportFormat.Csv);

            var data = new List<SampleData>()
            {
                new SampleData{ IntData=5, TextData="Hello"},
                new SampleData{ IntData=20, TextData="A,B;C"},
                new SampleData{ IntData=10, TextData="A\"B,C"},
            };

            var exportResult = exporter.Export(data);

            var reader = new StreamReader(exportResult);
            Assert.Equal(@"Number;Text", reader.ReadLine());
            Assert.Equal(@"5;Hello", reader.ReadLine());
            Assert.Equal(@"20;""A,B;C""", reader.ReadLine());
            Assert.Equal(@"10;""A""""B,C""", reader.ReadLine());

        }
    }

    public class SampleData
    {

        [ExportColumn("Number", 1)]
        public int IntData { get; set; }

        [ExportColumn("Text", 2)]
        public string TextData { get; set; }

    }
}
