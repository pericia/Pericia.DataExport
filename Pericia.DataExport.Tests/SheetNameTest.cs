using System.IO;
using Xunit;

namespace Pericia.DataExport
{
    public class SheetNameTest
    {
        [Fact]
        public void SheetsRenaming()
        {
            var exporter = new SheetNameTestExporter();
                        
            var sheet1 = exporter.TestSheetName(null);
            Assert.NotNull(sheet1);
            var sheet2 = exporter.TestSheetName(null);
            Assert.NotNull(sheet2);
            Assert.NotEqual(sheet1, sheet2);

            var sheet3 = exporter.TestSheetName("Te1st");
            Assert.Equal("Te1st", sheet3);
            var sheet4 = exporter.TestSheetName("Te1st");
            Assert.Equal("Te1st1", sheet4);
            var sheet5 = exporter.TestSheetName("Te1st");
            Assert.Equal("Te1st2", sheet5);
            var sheet6 = exporter.TestSheetName("Te1st1");
            Assert.Equal("Te1st3", sheet6);
        }



    }

    public class SheetNameTestExporter : XlsxDataExporter
    {
        public string TestSheetName(string? suggestedName)
        {
            return base.NewSheetName(suggestedName);
        }
    }
}
