using SBO.Hub.Attributes;
using SBO.Hub.Enums;
using SBO.Hub.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace SBO.Hub.Helpers
{
    public class TextFileHelper
    {
        public StreamWriter SW;
        public string Separator { get; set; } = String.Empty;
        public bool StartsWithSeparator { get; set; } = false;
        public PaddingTypeEnum NumericPaddingType { get; set; } = PaddingTypeEnum.Left;
        public PaddingTypeEnum AlphaNumericPaddingType { get; set; } = PaddingTypeEnum.Right;
        public string DecimalSeparator { get; set; } = "";
        public string DateFormat { get; set; } = "";
        public bool DecimalEmptyIfZero { get; set; } = false;

        private int writtenLinesQuantity = 0;
        public int WrittenLinesQuantity
        {
            get
            {
                return writtenLinesQuantity;
            }
        }

        public TextFileHelper()
        {
            SW = null;
        }

        public TextFileHelper(string filePath)
        {
            SW = new StreamWriter(filePath, false, Encoding.ASCII);
        }

        private int readLinesQuantity = 0;
        public int ReadLinesQuantity
        {
            get
            {
                return readLinesQuantity;
            }
        }

        public void WriteFile<T>(List<T> modelList, string fileName)
        {
            StreamWriter sw = new StreamWriter(fileName);
            try
            {
                foreach (var model in modelList)
                {
                    sw.WriteLine(WriteLine(model));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                sw.Close();
            }
        }

        public string WriteLine<T>(List<T> modelList)
        {
            if (SW == null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (var item in modelList)
                {
                    sb.AppendLine(this.WriteLine(item));
                }
                return sb.ToString();
            }
            else
            {
                foreach (var item in modelList)
                {
                    this.WriteLine(item);
                }
                return String.Empty;
            }
        }

        public string WriteLine(object model)
        {
            TextFileAttribute fileAttr;

            Type modelType = model.GetType();

            PropertyInfo[] propertiesList = modelType.GetProperties();
            List<FileAttributeModel> valuesList = new List<FileAttributeModel>();

            foreach (PropertyInfo property in propertiesList)
            {
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    fileAttr = attribute as TextFileAttribute;
                    if (fileAttr == null)
                    {
                        continue;
                    }
                    FileAttributeModel valueModel = new FileAttributeModel();
                    valueModel.Position = fileAttr.Position;

                    if (property.PropertyType == typeof(string) || property.PropertyType == typeof(String) ||
                        property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                    {
                        if (fileAttr.PaddingType == Enums.PaddingTypeEnum.NotSet)
                        {
                            fileAttr.PaddingType = AlphaNumericPaddingType;
                        }
                        if (String.IsNullOrEmpty(fileAttr.PaddingChar))
                        {
                            fileAttr.PaddingChar = " ";
                        }
                    }
                    else
                    {
                        if (fileAttr.PaddingType == Enums.PaddingTypeEnum.NotSet)
                        {
                            fileAttr.PaddingType = NumericPaddingType;
                        }
                        if (String.IsNullOrEmpty(fileAttr.PaddingChar))
                        {
                            fileAttr.PaddingChar = "0";
                        }
                    }

                    if (property.GetValue(model) == null)
                    {
                        valueModel.Value = String.Empty;
                    }
                    else
                    {
                        if (property.PropertyType == typeof(string) || property.PropertyType == typeof(String) ||
                            property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>))
                        {
                            valueModel.Value = property.GetValue(model).ToString();
                        }
                        else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<double>))
                        {
                            if (fileAttr.DecimalSeparator == null)
                            {
                                fileAttr.DecimalSeparator = DecimalSeparator;
                            }

                            valueModel.Value = Convert.ToDouble(property.GetValue(model)).ToString($"f{fileAttr.DecimalPlaces}").Replace(",", fileAttr.DecimalSeparator);
                            if (DecimalEmptyIfZero)
                            {
                                if (Convert.ToDouble(valueModel.Value) == 0)
                                {
                                    valueModel.Value = "";
                                }
                            }
                        }
                        else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                        {
                            if (String.IsNullOrEmpty(fileAttr.Format))
                            {
                                fileAttr.Format = DateFormat;
                            }
                            valueModel.Value = Convert.ToDateTime(property.GetValue(model)).ToString(fileAttr.Format);
                        }
                    }

                    if (fileAttr.OnylNumeric)
                    {
                        valueModel.Value = Regex.Replace(valueModel.Value, @"[^\d]", "");
                    }

                    if (fileAttr.Size > 0)
                    {
                        if (valueModel.Value.Length <= fileAttr.Size)
                        {
                            if (fileAttr.PaddingType == Enums.PaddingTypeEnum.Left)
                            {
                                valueModel.Value = valueModel.Value.PadLeft(fileAttr.Size, Convert.ToChar(fileAttr.PaddingChar));
                            }
                            else if (fileAttr.PaddingType == Enums.PaddingTypeEnum.Right)
                            {
                                valueModel.Value = valueModel.Value.PadRight(fileAttr.Size, Convert.ToChar(fileAttr.PaddingChar));
                            }
                        }
                        else
                        {
                            valueModel.Value = valueModel.Value.Substring(0, fileAttr.Size);
                        }
                    }
                    
                    valuesList.Add(valueModel);
                }
            }

            string line = String.Empty;
            if (StartsWithSeparator)
            {
                line += Separator;
            }
            valuesList = valuesList.OrderBy(v => v.Position).ToList();
            foreach (var valueModel in valuesList)
            {
                line += valueModel.Value + Separator;
            }

            writtenLinesQuantity++;
            if (SW != null)
            {
                SW.WriteLine(line);
            }
            return line;
        }

        public void CloseFile()
        {
            if (SW != null)
            {
                SW.Close();
            }
        }

        public List<T> ReadFile<T>(string fileName)
        {
            StreamReader sr = new StreamReader(fileName);
            List<T> list = new List<T>();
            try
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    list.Add(ReadLine<T>(line));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                sr.Close();
            }

            return list;
        }

        public T ReadLine<T>(string line)
        {
            var model = Activator.CreateInstance<T>();

            TextFileAttribute fileAttr;

            Type modelType = typeof(T);

            PropertyInfo[] propertiesList = modelType.GetProperties();

            string[] valuesList = null;
            if (!String.IsNullOrEmpty(Separator))
            {
                valuesList = line.Split(new string[] { Separator }, StringSplitOptions.None);
            }

            foreach (PropertyInfo property in propertiesList)
            {
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    fileAttr = attribute as TextFileAttribute;
                    if (fileAttr == null)
                    {
                        continue;
                    }

                    PropertyInfo modelProperty = model.GetType().GetProperty(property.Name);

                    string value = String.Empty;
                    if (String.IsNullOrEmpty(Separator))
                    {
                        value = line.Substring(fileAttr.Position, fileAttr.Size).Trim();
                    }
                    else
                    {
                        value = valuesList[fileAttr.Position];
                    }
                    
                    if (property.PropertyType == typeof(string) || property.PropertyType == typeof(String))
                    {
                        modelProperty.SetValue(model, value);
                    }
                    else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>))
                    {
                        int intValue;
                        if (Int32.TryParse(value, out intValue))
                        {
                            modelProperty.SetValue(model, intValue);
                        }
                    }
                    else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<double>))
                    {
                        if (String.IsNullOrEmpty(Separator))
                        {
                            string decimalPlace = line.Substring(fileAttr.Position + fileAttr.Size, fileAttr.DecimalPlaces);
                            double numericValue;
                            if (double.TryParse(value + "," + decimalPlace, out numericValue))
                            {
                                modelProperty.SetValue(model, numericValue);
                            }
                        }
                        else
                        {
                            NumberFormatInfo nfi = new NumberFormatInfo();
                            nfi.NumberDecimalSeparator = DecimalSeparator;
                            double numericValue;
                            if (double.TryParse(value, NumberStyles.AllowDecimalPoint, nfi, out numericValue))
                            {
                                modelProperty.SetValue(model, numericValue);
                            }
                        }
                    }
                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                    {
                        if (String.IsNullOrEmpty(fileAttr.Format))
                        {
                            fileAttr.Format = DateFormat;
                        }
                        DateTime dateValue;
                        if (DateTime.TryParseExact(value, fileAttr.Format, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateValue))
                        {
                            modelProperty.SetValue(model, dateValue);
                        }
                        else
                        {
                            throw new Exception($"Valor da posição {fileAttr.Position} não está no formato correto de data");
                        }
                    }
                }
            }
            readLinesQuantity++;
            return model;
        }
    }
}
