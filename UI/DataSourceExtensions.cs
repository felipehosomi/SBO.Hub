using SAPbouiCOM;
using SBO.Hub.Attributes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace SBO.Hub.UI
{
    public static class DataSourceExtensions
    {
        public static T FillModelFromDBDataSource<T>(this DBDataSource dataSource)
        {
            List<T> modelList = FillModelListFromDBDataSource<T>(dataSource);
            if (modelList.Count == 0)
            {
                return Activator.CreateInstance<T>();
            }
            else
            {
                return modelList[0];
            }
        }

        public static List<T> FillModelListFromDBDataSource<T>(this DBDataSource dataSource)
        {
            // Cria nova instância do model
            List<T> modelList = new List<T>();
            T model;

            HubModelAttribute hubModel;

            for (int i = 0; i < dataSource.Size; i++)
            {
                model = Activator.CreateInstance<T>();

                // Seta os valores no model
                foreach (PropertyInfo property in model.GetType().GetProperties())
                {
                    // Busca os Custom Attributes
                    foreach (Attribute attribute in property.GetCustomAttributes(true))
                    {
                        hubModel = attribute as HubModelAttribute;
                        if (hubModel != null)
                        {
                            if (!hubModel.AutoFill)
                            {
                                continue;
                            }

                            if (String.IsNullOrEmpty(hubModel.ColumnName))
                            {
                                hubModel.ColumnName = property.Name;
                            }

                            string value = dataSource.GetValue(hubModel.ColumnName, i);
                            if (String.IsNullOrEmpty(value))
                            {
                                continue;
                            }

                            if (property.PropertyType == typeof(string))
                            {
                                property.SetValue(model, value, null);
                            }
                            else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<double>))
                            {
                                property.SetValue(model, Convert.ToDouble(value.Replace(".", ",")), null);
                            }
                            else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(Nullable<decimal>))
                            {
                                property.SetValue(model, Convert.ToDecimal(value.Replace(".", ",")), null);
                            }
                            else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>))
                            {
                                property.SetValue(model, Convert.ToInt32(value.Replace(".", ",")), null);
                            }
                            else if (property.PropertyType == typeof(short) || property.PropertyType == typeof(Nullable<short>))
                            {
                                property.SetValue(model, Convert.ToInt16(value.Replace(".", ",")), null);
                            }
                            else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(Nullable<bool>))
                            {
                                property.SetValue(model, Convert.ToBoolean(value), null);
                            }
                            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                            {
                                property.SetValue(model, DateTime.ParseExact(value, "yyyyMMdd", CultureInfo.CurrentCulture), null);
                            }
                        }
                    }
                }
                modelList.Add(model);
            }
            return modelList;
        }
    }
}
