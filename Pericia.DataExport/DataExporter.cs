using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Pericia.DataExport
{
    public abstract class DataExporter
    {
        protected MemoryStream stream = new MemoryStream();

        public void AddSheet<T>(IEnumerable<T> data, string name=null)
        {
            NewSheet(name);

            var typeInfo = typeof(T).GetTypeInfo();

            var properties = new List<ColumnInfo>();
            foreach (var prop in typeInfo.DeclaredProperties)
            {
                var attribute = prop.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    properties.Add(new ColumnInfo() { Prop = prop, Attr = attribute });
                }
            }
            
            properties = properties.OrderBy(a => a.Attr.Order).ToList();
            // Write headers
            foreach (var prop in properties)
            {
                WriteData(prop.Attr.Title);
            }
            NewLine();

            foreach (var line in data)
            {
                foreach (var prop in properties)
                {
                    WriteData(prop.Prop.GetValue(line).ToString());
                }
                NewLine();
            }
        }


        public MemoryStream Export<T>(IEnumerable<T> data)
        {
            AddSheet(data);

            return GetFile();
        }

#if !NETSTANDARD1_3
        public void AddSheet(System.Data.Common.DbDataReader reader, string name = null)
        {
            NewSheet(name);

            // Write headers
            for (int i = 0; i < reader.FieldCount; i++)
            {
                WriteData(reader.GetName(i));
            }
            NewLine();

            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    WriteData(reader.GetValue(i).ToString());
                }
                NewLine();
            }
        }

        public MemoryStream Export(System.Data.Common.DbDataReader reader)
        {
            AddSheet(reader);

            return GetFile();
        }

#endif

        public abstract MemoryStream GetFile();



        protected abstract void NewSheet(string name);
        protected abstract void NewLine();
        protected abstract void WriteData(string data);

        private class ColumnInfo
        {
            internal PropertyInfo Prop { get; set; }
            internal ExportColumnAttribute Attr { get; set; }
        }
    }
}
