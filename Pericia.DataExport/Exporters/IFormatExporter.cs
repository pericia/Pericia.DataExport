using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Pericia.DataExport.Exporters
{
    internal interface IFormatExporter
    {
        void WriteData(string data);
        void NewLine();
        void NewSheet();

        Stream GetStream();
    }
}
