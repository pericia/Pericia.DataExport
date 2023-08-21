using DocumentFormat.OpenXml.Math;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Pericia.DataExport
{
    public abstract class DataExporter : IDisposable
    {
        protected MemoryStream stream { get; } = new MemoryStream();

        private Dictionary<Type, Func<object, object>>? TypeDataConverter { get; set; }
        

        public void AddSheet(IEnumerable<object> data, string[] columns, string? name = null)
        {
            var columnsObject = columns.Select(c => new ExportColumn { Property = c, Title = c });
            AddSheet(data, columnsObject.ToArray(), name);
        }

        public void AddSheet(IEnumerable<object> data, ExportColumn[] columns, string? name = null)
        {
            NewSheet(name);

            // Write headers
            foreach (var column in columns)
            {
                WriteDataRaw(column.Title);
            }
            NewLine();

            foreach (var line in data)
            {
                var lineType = line.GetType();

                foreach (var column in columns)
                {
                    var property = lineType.GetProperty(column.Property, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
                    var value = property?.GetValue(line);
                    WriteDataValue(value ?? "");
                }

                NewLine();
            }
        }

        public void AddSheet<T>(IEnumerable<T> data, string? name = null)
        {
            NewSheet(name);

            var typeInfo = typeof(T).GetTypeInfo();

            var properties = new List<ColumnInfo>();
            foreach (var prop in typeInfo.DeclaredProperties)
            {
                var attribute = prop.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    properties.Add(new ColumnInfo(prop, attribute));
                }
            }

            properties = properties.OrderBy(a => a.Attr.Order).ToList();
            // Write headers
            foreach (var prop in properties)
            {
                WriteDataRaw(prop.Attr.Title);
            }
            NewLine();

            foreach (var line in data)
            {
                foreach (var prop in properties)
                {
                    WriteDataValue(prop.Prop.GetValue(line));
                }
                NewLine();
            }
        }

        public void AddSheet(System.Data.Common.DbDataReader reader, string? name = null)
        {
            NewSheet(name);

            // Write headers
            for (int i = 0; i < reader.FieldCount; i++)
            {
                WriteDataRaw(reader.GetName(i));
            }
            NewLine();

            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var value = reader.GetValue(i);
                    if (value == null || value == DBNull.Value)
                    {
                        value = "";
                    }
                    WriteDataValue(value);
                }
                NewLine();
            }
        }

        public MemoryStream Export<T>(IEnumerable<T> data)
        {
            AddSheet(data);

            return GetFile();
        }

        public MemoryStream Export(System.Data.Common.DbDataReader reader)
        {
            AddSheet(reader);

            return GetFile();
        }

        public abstract MemoryStream GetFile();


        protected abstract void NewSheet(string? name);
        protected abstract void NewLine();

        public void AddDataConverter<T>(Func<T, object> converter)
        {
            if (TypeDataConverter == null)
            {
                TypeDataConverter = new Dictionary<Type, Func<object, object>>();
            }

            TypeDataConverter.Add(typeof(T), o => converter((T)o));
        }


        private void WriteDataValue(object data)
        {
            if (TypeDataConverter != null)
            {
                var type = data.GetType();
                if (TypeDataConverter.ContainsKey(type))
                {
                    var converter = TypeDataConverter[type];
                    data = converter(data);
                }
            }

            WriteDataRaw(data);
        }

        protected abstract void WriteDataRaw(object data);

        private class ColumnInfo
        {
            public ColumnInfo(PropertyInfo prop, ExportColumnAttribute attr)
            {
                Prop = prop;
                Attr = attr;
            }

            internal PropertyInfo Prop { get; set; }
            internal ExportColumnAttribute Attr { get; set; }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            stream.Dispose();
        }
    }
}
