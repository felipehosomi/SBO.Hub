using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using System.Reflection;
using System.Runtime.InteropServices;
using SBO.Hub.Attributes;

namespace SBO.Hub.UI
{
    public static class DataTableExtensions
    {
        /// <summary>
        /// Preenche tabela com a lista
        /// </summary>
        /// <typeparam name="T">Tipo da lista</typeparam>
        /// <param name="table">Tabela</param>
        /// <param name="modelList">Lista</param>
        /// <param name="stopOnError">Para caso algum campo do model não exista na tabela</param>
        /// <returns></returns>
        public static DataTable CreateColumns(this DataTable table, Type modelType)
        {
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    HubModelAttribute hubModel = attribute as HubModelAttribute;
                    if (hubModel != null)
                    {
                        if (hubModel.UIIgnore)
                        {
                            continue;
                        }
                        if (String.IsNullOrEmpty(hubModel.UIFieldName))
                        {
                            hubModel.UIFieldName = property.Name;
                        }
                        try
                        {
                            BoFieldsType columnType;

                            if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<decimal>) || property.PropertyType == typeof(Nullable<double>))
                            {
                                columnType = BoFieldsType.ft_Float;
                            }
                            else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>))
                            {
                                columnType = BoFieldsType.ft_Integer;
                            }
                            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                            {
                                columnType = BoFieldsType.ft_Date;
                            }
                            else
                            {
                                columnType = BoFieldsType.ft_Text;
                            }

                            table.Columns.Add(hubModel.UIFieldName, columnType);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
            }

            return table;
        }

        /// <summary>
        /// Preenche tabela com a lista
        /// </summary>
        /// <typeparam name="T">Tipo da lista</typeparam>
        /// <param name="table">Tabela</param>
        /// <param name="modelList">Lista</param>
        /// <param name="stopOnError">Para caso algum campo do model não exista na tabela</param>
        /// <returns></returns>
        public static DataTable FillTable<T>(this DataTable table, List<T> modelList, bool stopOnError = false)
        {
            table.Rows.Add(modelList.Count);

            Type modelType = typeof(T);
            // Seta os valores no model
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    int i = 0;
                    foreach (var item in modelList)
                    {
                        HubModelAttribute hubModel = attribute as HubModelAttribute;
                        if (hubModel != null)
                        {
                            if (hubModel.UIIgnore)
                            {
                                continue;
                            }
                            if (String.IsNullOrEmpty(hubModel.UIFieldName))
                            {
                                hubModel.UIFieldName = property.Name;
                            }
                            try
                            {
                                table.SetValue(hubModel.UIFieldName, i, property.GetValue(item, null));
                            }
                            catch (Exception ex)
                            {
                                if (stopOnError)
                                {
                                    throw ex;
                                }
                            }
                        }
                        i++;
                    }
                }
            }
            return table;
        }

        /// <summary>
        /// Preenche model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="table">Tabela</param>
        /// <param name="showProgressBar">Exibir progress bar?</param>
        /// <returns></returns>
        public static T FillModel<T>(this DataTable table, int row)
        {
            // Cria nova instância do model
            T model = Activator.CreateInstance<T>();
            if (row < 0)
            {
                return model;
            }

            HubModelAttribute hubModel;

            // Seta os valores no model
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;
                    if (hubModel != null)
                    {
                        if (hubModel.UIIgnore)
                        {
                            continue;
                        }
                        if (String.IsNullOrEmpty(hubModel.UIFieldName))
                        {
                            break;
                        }
                        else
                        {
                            property.SetValue(model, table.Columns.Item(hubModel.UIFieldName).Cells.Item(row).Value, null);
                        }
                    }
                }
            }
            return model;
        }

        /// <summary>
        /// Preenche model de acordo com os nomes das propriedades, sem necessidade do HubModelAttribute
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="table">Tabela</param>
        /// <returns></returns>
        public static List<T> FillByModelProperties<T>(this DataTable table)
        {
            List<T> modelList = new List<T>();
            // Cria nova instância do model
            T model;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                model = Activator.CreateInstance<T>();
                foreach (PropertyInfo property in model.GetType().GetProperties())
                {
                    try
                    {
                        property.SetValue(model, table.GetValue(property.Name, i), null);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Erro ao setar campo {property.Name}: {ex.Message}");
                    }
                }
                modelList.Add(model);
            }

            return modelList;
        }

        /// <summary>
        /// Preenche model de acordo com o valor de uma coluna
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="table">Tabela</param>
        /// <param name="showProgressBar">Exibir progress bar?</param>
        /// <param name="columnToCheck">Nome da coluna para validar</param>
        /// <param name="valueToCheck">Valor a ser validado</param>
        /// <param name="stopOnError">Para caso algum campo da tabela que não exista no model</param>
        /// <returns></returns>
        public static List<T> FillModelByColumnValue<T>(this DataTable table, bool showProgressBar = false, string columnToCheck = "#", string valueToCheck = "Y", bool stopOnError = false)
        {
            List<T> modelList = new List<T>();
            // Cria nova instância do model
            T model;

            ProgressBar pgb = null;
            if (showProgressBar)
            {
                pgb = SBOApp.Application.StatusBar.CreateProgressBar("Carregando dados da tela", table.Rows.Count, false);
            }
            HubModelAttribute hubModel;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (showProgressBar)
                {
                    pgb.Value++;
                }

                if (table.GetValue(columnToCheck, i).ToString() != valueToCheck.ToString())
                {
                    continue;
                }

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
                            if (hubModel.UIIgnore)
                            {
                                continue;
                            }
                            if (String.IsNullOrEmpty(hubModel.UIFieldName))
                            {
                                hubModel.UIFieldName = property.Name;
                            }

                            try
                            {
                                property.SetValue(model, table.GetValue(hubModel.UIFieldName, i), null);
                            }
                            catch
                            {
                                try
                                {
                                    if (property.PropertyType == typeof(string))
                                    {
                                        property.SetValue(model, table.GetValue(hubModel.UIFieldName, i).ToString(), null);
                                    }
                                    else if (property.PropertyType == typeof(int))
                                    {
                                        property.SetValue(model, Convert.ToInt32(table.GetValue(hubModel.UIFieldName, i).ToString()), null);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (stopOnError)
                                    {
                                        throw ex;
                                    }
                                }
                            }
                        }
                    }
                }
                modelList.Add(model);
            }
            if (pgb != null)
            {
                pgb.Stop();
                Marshal.ReleaseComObject(pgb);
                pgb = null;
            }

            return modelList;
        }

        /// <summary>
        /// Preenche model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="table">Tabela</param>
        /// <param name="showProgressBar">Exibir progress bar?</param>
        /// <returns></returns>
        public static List<T> FillModel<T>(this DataTable table, bool showProgressBar = false)
        {
            List<T> modelList = new List<T>();
            // Cria nova instância do model
            T model;

            ProgressBar pgb = null;
            if (showProgressBar)
            {
                pgb = SBOApp.Application.StatusBar.CreateProgressBar("Carregando dados da tela", table.Rows.Count, false);
            }
            HubModelAttribute hubModel;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (showProgressBar)
                {
                    pgb.Value++;
                }
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
                            if (hubModel.UIIgnore)
                            {
                                continue;
                            }
                            if (String.IsNullOrEmpty(hubModel.UIFieldName))
                            {
                                hubModel.UIFieldName = hubModel.Description;
                            }

                            if (String.IsNullOrEmpty(hubModel.UIFieldName))
                            {
                                break;
                            }
                            else
                            {
                                property.SetValue(model, table.Columns.Item(hubModel.UIFieldName).Cells.Item(i).Value, null);
                            }
                        }
                    }
                }
                modelList.Add(model);
            }
            if (pgb != null)
            {
                pgb.Stop();
                Marshal.ReleaseComObject(pgb);
                pgb = null;
            }

            return modelList;
        }

        /// <summary>
        /// Preenche lista de acordo com valor em determinada coluna
        /// </summary>
        /// <param name="columnName">Nome da coluna que irá retornar na lista</param>
        /// <param name="table">Tabela</param>
        /// <returns></returns>
        public static List<string> FillStringListByColumnValue(this DataTable table, string columnName)
        {
            List<string> list = new List<string>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                list.Add(table.GetValue(columnName, i).ToString());
            }

            return list;
        }

        /// <summary>
        /// Preenche lista de acordo com valor em determinada coluna
        /// </summary>
        /// <param name="columnName">Nome da coluna que irá retornar na lista</param>
        /// <param name="columnToCheck">Nome da coluna para verificar o valor</param>
        /// <param name="valueToCheck">Valor a ser verificado</param>
        /// <param name="table">Tabela</param>
        /// <returns></returns>
        public static List<string> FillStringListByColumnValue(this DataTable table, string columnName, string columnToCheck = "#", string valueToCheck = "Y")
        {
            List<string> list = new List<string>();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.GetValue(columnToCheck, i).ToString() == valueToCheck)
                {
                    list.Add(table.GetValue(columnName, i).ToString());
                }
            }

            return list;
        }
    }
}
