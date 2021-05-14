using SAPbouiCOM;
using SBO.Hub.Attributes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace SBO.Hub.UI
{
    public static class MatrixExtensions
    {
        #region AddMatrixRow
        public static Matrix AddMatrixRow(this Matrix matrix, IDBDataSource dBDataSource, bool focusCell = false)
        {
            matrix.FlushToDataSource();
            // Insere um novo registro vazio dentro do data source
            dBDataSource.InsertRecord(dBDataSource.Size);

            if (dBDataSource.Size == 1)
            {
                dBDataSource.InsertRecord(dBDataSource.Size);
            }

            if (matrix.VisualRowCount.Equals(0))
            {
                dBDataSource.RemoveRecord(0);
            }

            // Loads the user interface with current data from the matrix objects data source.
            matrix.LoadFromDataSourceEx(false);

            if (focusCell)
                matrix.SetCellFocus(matrix.VisualRowCount, 1);
            return matrix;
        }
        #endregion

        #region RemoveMatrixRow
        public static Matrix RemoveMatrixRow(this Matrix matrix, IDBDataSource dBDataSource, int row)
        {
            dBDataSource.RemoveRecord(row - 1);

            // Loads the user interface with current data from the matrix objects data source.
            matrix.LoadFromDataSource();
            return matrix;
        }
        #endregion

        public static int GetColumnIndex(this Matrix mtx, string columnId)
        {
            //get the matrix as a xml
            String matrixSchemaXML = mtx.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
            //find your UniqueId "<UniqueID>Your ID</UniqueID>"
            int indexOfText = matrixSchemaXML.IndexOf($"<UniqueID>{columnId}</UniqueID>");
            //get the substring
            string subs = matrixSchemaXML.Substring(0, indexOfText);
            //count a number of occurences <ColumnInfo> - 1, because matrix column starts in index 0
            int index = Regex.Matches(subs, "<ColumnInfo>").Count - 1;
            return index;
        }

        public static T FillModelFromMatrix<T>(this Matrix mtx)
        {
            List<T> modelList = FillModelListFromMatrix<T>(mtx);
            if (modelList.Count == 0)
            {
                return Activator.CreateInstance<T>();
            }
            else
            {
                return modelList[0];
            }
        }

        public static List<T> FillModelListFromMatrix<T>(this Matrix mtx, string columnToCheck = "", string valueToCheck = "Y", SBO.Hub.Enums.UIFieldTypeEnum columnTypeToCheck = Enums.UIFieldTypeEnum.CheckBox)
        {
            // Cria nova instância do model
            List<T> modelList = new List<T>();
            T model;

            HubModelAttribute hubModel;

            for (int i = 1; i < mtx.RowCount + 1; i++)
            {
                string currentValue = "";
                if (!String.IsNullOrEmpty(columnToCheck))
                {
                    switch (columnTypeToCheck)
                    {
                        case SBO.Hub.Enums.UIFieldTypeEnum.EditText:
                            EditText etx = (EditText)mtx.Columns.Item(columnToCheck).Cells.Item(i).Specific;
                            currentValue = etx.Value;
                            break;
                        case SBO.Hub.Enums.UIFieldTypeEnum.PriceEditText:
                            EditText etxPrice = (EditText)mtx.Columns.Item(columnToCheck).Cells.Item(i).Specific;
                            currentValue = etxPrice.Value.Split(' ')[0];
                            break;
                        case SBO.Hub.Enums.UIFieldTypeEnum.ComboBox:
                            ComboBox comboBox = (ComboBox)mtx.Columns.Item(columnToCheck).Cells.Item(i).Specific;
                            currentValue = comboBox.Value;
                            break;
                        case SBO.Hub.Enums.UIFieldTypeEnum.CheckBox:
                            CheckBox checkBox = (CheckBox)mtx.Columns.Item(columnToCheck).Cells.Item(i).Specific;
                            if (checkBox.Checked)
                            {
                                currentValue = checkBox.ValOn;
                            }
                            else
                            {
                                currentValue = checkBox.ValOff;
                            }
                            break;
                        default:
                            break;
                    }
                    if (currentValue != valueToCheck)
                    {
                        continue;
                    }
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
                            string value = String.Empty;

                            switch (hubModel.UIFieldType)
                            {
                                case SBO.Hub.Enums.UIFieldTypeEnum.EditText:
                                    EditText etx = (EditText)mtx.Columns.Item(hubModel.UIFieldName).Cells.Item(i).Specific;
                                    value = etx.Value;
                                    break;
                                case SBO.Hub.Enums.UIFieldTypeEnum.PriceEditText:
                                    EditText etxPrice = (EditText)mtx.Columns.Item(hubModel.UIFieldName).Cells.Item(i).Specific;
                                    value = etxPrice.Value.Split(' ')[0];
                                    break;
                                case SBO.Hub.Enums.UIFieldTypeEnum.ComboBox:
                                    ComboBox comboBox = (ComboBox)mtx.Columns.Item(hubModel.UIFieldName).Cells.Item(i).Specific;
                                    value = comboBox.Value;
                                    break;
                                case SBO.Hub.Enums.UIFieldTypeEnum.CheckBox:
                                    CheckBox checkBox = (CheckBox)mtx.Columns.Item(hubModel.UIFieldName).Cells.Item(i).Specific;
                                    if (checkBox.Checked)
                                    {
                                        value = checkBox.ValOn;
                                    }
                                    else
                                    {
                                        value = checkBox.ValOff;
                                    }
                                    break;
                                default:
                                    break;
                            }

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
