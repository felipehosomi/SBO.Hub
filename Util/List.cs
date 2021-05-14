using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBO.Hub.Util
{
    public static class List
    {
        public static DataTable ToDataTable<T>(this IList<T> list) where T : class
        {
            try
            {
                var table = CreateDataTable<T>();
                var objType = typeof(T);
                var properties = TypeDescriptor.GetProperties(objType);
                foreach (var item in list)
                {
                    var row = table.NewRow();
                    foreach (PropertyDescriptor property in properties)
                    {
                        if (!CanUseType(property.PropertyType)) continue;
                        row[property.Name] = property.GetValue(item) ?? DBNull.Value;
                    }
                    table.Rows.Add(row);
                }
                return table;
            }
            catch (DataException)
            {
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static DataTable CreateDataTable<T>() where T : class
        {
            var objType = typeof(T);
            var table = new DataTable(objType.Name);
            var properties = TypeDescriptor.GetProperties(objType);
            foreach (PropertyDescriptor property in properties)
            {
                var propertyType = property.PropertyType;
                if (!CanUseType(propertyType)) continue;

                //nullables must use underlying types
                if (propertyType.IsGenericType && (propertyType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                    propertyType = Nullable.GetUnderlyingType(propertyType);
                //enums also need special treatment
                if (propertyType.IsEnum)
                    propertyType = Enum.GetUnderlyingType(propertyType);
                table.Columns.Add(property.Name, propertyType);
            }
            return table;
        }

        private static bool CanUseType(Type propertyType)
        {
            //only strings and value types
            if (propertyType.IsArray) return false;
            if (!propertyType.IsValueType && (propertyType != typeof(string))) return false;
            return true;
        }
    }
}
