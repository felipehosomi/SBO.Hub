using SAPbobsCOM;
using SBO.Hub.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace SBO.Hub.DAO
{
    public class UserChildObjectDAO
    {
        public string ParentTableName { get; set; }
        public string TableName { get; set; }
        public object Model { get; set; }

        public List<object> ModelList { get; set; }
        public BoUTBTableType UserTableType { get; set; }

        public UserChildObjectDAO(string parentTableName, string tableName)
        {
            ParentTableName = parentTableName;
            TableName = tableName;
            UserTableType = BoUTBTableType.bott_MasterDataLines;
        }

        public void CreateModel(object parentCode)
        {
            if (Model == null)
            {
                throw new Exception("Informe o model a ser criado!");
            }

            this.ModelList = new List<object>();
            this.ModelList.Add(Model);
            this.CreateModelList(parentCode);
        }

        public void CreateModelList(object parentCode)
        {
            if (ModelList == null || ModelList.Count == 0)
            {
                return;
            }
            CompanyService oCompanyService = SBOApp.Company.GetCompanyService();
            GeneralService oGeneralService = oCompanyService.GetGeneralService(ParentTableName.Replace("@", ""));

            GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            if (UserTableType == BoUTBTableType.bott_MasterDataLines)
            {
                oGeneralParams.SetProperty("Code", parentCode);
            }
            else
            {
                oGeneralParams.SetProperty("DocEntry", parentCode);
            }

            GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);
            GeneralDataCollection oChildren = oGeneralData.Child(TableName.Replace("@", ""));
            GeneralData oChild = null;

            try
            {
                foreach (object model in ModelList)
                {
                    Model = model;
                    oChild = oChildren.Add();

                    HubModelAttribute hubModel;
                    object value;
                    // Percorre as propriedades do Model
                    foreach (PropertyInfo property in Model.GetType().GetProperties())
                    {
                        // Busca os Custom Attributes
                        foreach (Attribute attribute in property.GetCustomAttributes(true))
                        {
                            hubModel = attribute as HubModelAttribute;
                            if (property.GetType() != typeof(DateTime))
                                value = property.GetValue(Model, null);
                            else
                                value = ((DateTime)property.GetValue(Model, null)).ToString("yyyy-MM-dd HH:mm:ss");

                            if (hubModel != null)
                            {
                                // Se não for DataBaseField não seta nas properties
                                if (!hubModel.DataBaseFieldYN)
                                {
                                    break;
                                }
                                if (String.IsNullOrEmpty(hubModel.ColumnName))
                                {
                                    hubModel.ColumnName = property.Name;
                                }
                                if (hubModel.ColumnName == "Code" || hubModel.ColumnName == "LineId")
                                {
                                    continue;
                                }

                                if (value == null)
                                {
                                    if (property.PropertyType == typeof(string))
                                    {
                                        value = String.Empty;
                                    }
                                    else if (property.PropertyType != typeof(DateTime) && property.PropertyType != typeof(Nullable<DateTime>))
                                    {
                                        value = 0;
                                    }
                                    else
                                    {
                                        value = new DateTime();
                                    }
                                }

                                if (property.PropertyType != typeof(decimal) && property.PropertyType != typeof(Nullable<decimal>))
                                {
                                    oChild.SetProperty(hubModel.ColumnName, value);
                                }
                                else
                                {
                                    oChild.SetProperty(hubModel.ColumnName, Convert.ToDouble(value));
                                }
                                break;
                            }
                        }
                    }
                }

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Marshal.ReleaseComObject(oGeneralService);
                Marshal.ReleaseComObject(oGeneralData);
                Marshal.ReleaseComObject(oGeneralParams);
                Marshal.ReleaseComObject(oCompanyService);

                oGeneralService = null;
                oGeneralData = null;
                oCompanyService = null;
                oGeneralParams = null;

                if (oChildren != null)
                {
                    Marshal.ReleaseComObject(oChildren);
                    oChildren = null;
                }
                if (oChild != null)
                {
                    Marshal.ReleaseComObject(oChild);
                    oChild = null;
                }
            }
        }

        public void UpdateModel(object parentCode)
        {
            if (Model == null)
            {
                throw new Exception("Informe o model a ser criado!");
            }

            this.ModelList = new List<object>();
            this.ModelList.Add(Model);
            this.UpdateModelList(parentCode);
        }

        public void UpdateModelList(object parentCode)
        {
            Recordset rstExistsParent = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = "SELECT TOP 1 1 FROM [{0}] WHERE Code = '{1}'";
            sql = String.Format(sql, ParentTableName, parentCode);

            try
            {
                rstExistsParent.DoQuery(sql);

                if (rstExistsParent.RecordCount == 0)
                {
                    return;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(rstExistsParent);
                rstExistsParent = null;
            }

            CompanyService oCompanyService = SBOApp.Company.GetCompanyService();
            GeneralService oGeneralService = oCompanyService.GetGeneralService(ParentTableName.Replace("@", ""));

            GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("Code", parentCode);

            GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);
            GeneralDataCollection oChildren = oGeneralData.Child(TableName.Replace("@", ""));
            GeneralData oChild = null;

            try
            {
                int lineId;
                foreach (object model in ModelList)
                {
                    Model = model;
                    lineId = Convert.ToInt32(model.GetType().GetProperty("LineId").GetValue(Model, null)) - 1;
                    oChild = oChildren.Item(lineId);

                    HubModelAttribute hubModel;
                    object value;
                    // Percorre as propriedades do Model
                    foreach (PropertyInfo property in Model.GetType().GetProperties())
                    {
                        // Busca os Custom Attributes
                        foreach (Attribute attribute in property.GetCustomAttributes(true))
                        {
                            hubModel = attribute as HubModelAttribute;
                            if (property.GetType() != typeof(DateTime))
                                value = property.GetValue(Model, null);
                            else
                                value = ((DateTime)property.GetValue(Model, null)).ToString("yyyy-MM-dd HH:mm:ss");

                            if (hubModel != null)
                            {
                                // Se não for DataBaseField não seta nas properties
                                if (!hubModel.DataBaseFieldYN)
                                {
                                    break;
                                }
                                
                                if (String.IsNullOrEmpty(hubModel.ColumnName))
                                {
                                    hubModel.ColumnName = property.Name;
                                }
                                if (hubModel.ColumnName == "Code" || hubModel.ColumnName == "LineId")
                                {
                                    continue;
                                }

                                if (value == null)
                                {
                                    if (property.PropertyType == typeof(string))
                                    {
                                        value = String.Empty;
                                    }
                                    else if (property.PropertyType != typeof(DateTime) && property.PropertyType != typeof(Nullable<DateTime>))
                                    {
                                        value = 0;
                                    }
                                    else
                                    {
                                        value = new DateTime();
                                    }
                                }

                                if (property.PropertyType != typeof(decimal) && property.PropertyType != typeof(Nullable<decimal>))
                                {
                                    oChild.SetProperty(hubModel.ColumnName, value);
                                }
                                else
                                {
                                    oChild.SetProperty(hubModel.ColumnName, Convert.ToDouble(value));
                                }
                                break;
                            }
                        }
                    }
                }

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Marshal.ReleaseComObject(oGeneralService);
                Marshal.ReleaseComObject(oGeneralData);
                Marshal.ReleaseComObject(oCompanyService);

                oGeneralService = null;
                oGeneralData = null;
                oCompanyService = null;

                if (oChildren != null)
                {
                    Marshal.ReleaseComObject(oChildren);
                    oChildren = null;
                }
                if (oChild != null)
                {
                    Marshal.ReleaseComObject(oChild);
                    oChild = null;
                }
            }
        }
    }
}
