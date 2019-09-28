using System;
using System.Collections.Generic;
using System.Text;

namespace Pericia.DataExport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportColumnAttribute : Attribute
    {
        public string Title { get; set; } = string.Empty;

        public int Order { get; set; }
    }
}
