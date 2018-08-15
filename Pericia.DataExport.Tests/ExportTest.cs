using System;
using Xunit;

namespace Pericia.DataExport.Tests
{
    public class ExportTest
    {
        [Fact]
        public void Test1()
        {
            var exporter = new DataExporter();

            exporter.Export("");
        }
    }
}
