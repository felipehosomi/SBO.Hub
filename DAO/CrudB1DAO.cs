using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using SAPbobsCOM;
using SBO.Hub.Attributes;
using SBO.Hub.Enums;
using SBO.Hub.Util;

namespace SBO.Hub.DAO
{
    public class CrudB1DAO
    {
        public string TableName { get; set; }
        public object Model { get; set; }
        public BoUTBTableType UserTableType { get; set; }

        public CrudB1DAO()
        {
            UserTableType = BoUTBTableType.bott_MasterData;
        }

        public CrudB1DAO(string tableName)
        {
            TableName = tableName;
            UserTableType = BoUTBTableType.bott_MasterData;
        }

        public void BeginTransaction()
        {
            if (!SBOApp.Company.InTransaction)
            {
                SBOApp.Company.StartTransaction();
            }
        }

        public void CommitTransaction()
        {
            if (SBOApp.Company.InTransaction)
            {
                SBOApp.Company.EndTransaction(BoWfTransOpt.wf_Commit);
            }
        }

        public void RollbackTransaction()
        {
            if (SBOApp.Company.InTransaction)
            {
                SBOApp.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
            }
        }

        #region CRUD
        #region CreateUpdateModel
        /// <summary>
        /// Insere dados no banco
        /// </summary>
        /// <param name="model">Objeto do tipo Model</param>
        /// <param name="tableName">Nome da tabela</param>
        public string CreateModel()
        {
            switch (UserTableType)
            {
                case BoUTBTableType.bott_NoObject:
                case BoUTBTableType.bott_NoObjectAutoIncrement:
                    return this.SaveNonObjectModel(EnumCrudOperation.Create);
                default:
                    return this.SaveModel(EnumCrudOperation.Create);
            }
        }

        public void DeleteModel()
        {
            switch (UserTableType)
            {
                case BoUTBTableType.bott_NoObject:
                    this.SaveNonObjectModel(EnumCrudOperation.Delete);
                    break;
            }
        }

        /// <summary>
        /// Atualiza dados no banco
        /// </summary>
        /// <param name="model">Objeto do tipo Model</param>
        /// <param name="tableName">Nome da tabela</param>
        public void UpdateModel()
        {
            switch (UserTableType)
            {
                case BoUTBTableType.bott_NoObject:
                    this.SaveNonObjectModel(EnumCrudOperation.Update);
                    break;
                default:
                    this.SaveModel(EnumCrudOperation.Update);
                    break;
            }
        }

        /// <summary>
        /// Atualiza dados no banco
        /// </summary>
        /// <param name="model">Objeto do tipo Model</param>
        /// <param name="tableName">Nome da tabela</param>
        private void UpdateModel(object whereModel)
        {
            HubModelAttribute hubModel;
            StringBuilder where = new StringBuilder();

            object value;
            foreach (PropertyInfo property in whereModel.GetType().GetProperties())
            {
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;
                    if (hubModel != null)
                    {
                        value = property.GetValue(whereModel, null);
                        if (String.IsNullOrEmpty(hubModel.ColumnName))
                            hubModel.ColumnName = property.Name;
                        if (value != null)
                        {
                            switch (value.GetType().ToString())
                            {
                                case "String":
                                    if (!String.IsNullOrEmpty(value.ToString()))
                                        where.AppendFormat("AND {0} = '{1}' ", hubModel.ColumnName, value);
                                    break;
                                case "int":
                                case "double":
                                case "decimal":
                                    where.AppendFormat("AND {0} = {1} ", hubModel.ColumnName, value);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            this.UpdateModel(where.ToString());
        }

        /// <summary>
        /// Atualiza dados no banco
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição WHRE</param>
        /// <param name="model">Model com os dados a serem atualizados</param>
        public void UpdateModel(string where)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            string sql = @"SELECT DocEntry FROM [{0}] WHERE {1}";
            sql = SBOApp.TranslateToHana(String.Format(sql, TableName, where));

            rs.DoQuery(String.Format(sql, TableName, where));
            if (rs.RecordCount > 0)
            {
                CompanyService sCompany = SBOApp.Company.GetCompanyService();
                GeneralService oGeneralService = sCompany.GetGeneralService(TableName.Replace("@", ""));

                GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", rs.Fields.Item(0).Value.ToString());

                GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                HubModelAttribute hubModel;
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    foreach (Attribute attribute in property.GetCustomAttributes(true))
                    {
                        hubModel = attribute as HubModelAttribute;
                        if (hubModel.DataBaseFieldYN)
                        {
                            if (String.IsNullOrEmpty(hubModel.ColumnName))
                            {
                                hubModel.ColumnName = property.Name;
                            }
                            oGeneralData.SetProperty(hubModel.ColumnName, property.GetValue(Model, null));
                        }
                    }
                }
                oGeneralService.Update(oGeneralData);
            }

            Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();
        }

        /// <summary>
        /// Salva o model no BD de acordo com o tipo da operação
        /// </summary>
        /// <param name="enumCrudOperation">Operação - Create ou Update</param>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="model">Modelo</param>
        private string SaveModel(EnumCrudOperation enumCrudOperation)
        {
            CompanyService sCompany = null;
            GeneralService oGeneralService = null;
            GeneralData oGeneralData = null;


            object code = null;

            try
            {
                sCompany = SBOApp.Company.GetCompanyService();
                oGeneralService = sCompany.GetGeneralService(TableName.Replace("@", ""));
                oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);


                if (enumCrudOperation == EnumCrudOperation.Update)
                {
                    GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);

                    try
                    {
                        code = Model.GetType().GetProperty("Code").GetValue(Model, null);
                    }
                    catch
                    {
                        Recordset rstDocEntry = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string sql = @"SELECT DocEntry FROM {0} WHERE Code = '{1}'";
                        sql = String.Format(sql, TableName, Model.GetType().GetProperty("Code").GetValue(Model, null));
                        sql = SBOApp.TranslateToHana(sql);

                        rstDocEntry.DoQuery(sql);
                        if (rstDocEntry.RecordCount > 0)
                        {
                            code = rstDocEntry.Fields.Item(0).Value;
                        }

                        Marshal.ReleaseComObject(oGeneralParams);
                        oGeneralParams = null;

                        Marshal.ReleaseComObject(rstDocEntry);
                        rstDocEntry = null;
                    }

                    oGeneralParams.SetProperty("Code", code);
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                }
                else
                {
                    try
                    {
                        code = Model.GetType().GetProperty("Code").GetValue(Model, null);
                    }
                    catch (Exception ex)
                    {






                        code = GetNextCode(TableName).PadLeft(10, '0');
                        oGeneralData.SetProperty("Code", code);

                    }
                }

                HubModelAttribute hubModel;
                object value;
                // Percorre as propriedades do Model
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    try
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
                                // Se não for DataBaseField ou for readonly não seta nas properties
                                if (!hubModel.DataBaseFieldYN)
                                {
                                    break;
                                }
                                if (String.IsNullOrEmpty(hubModel.ColumnName))
                                {
                                    hubModel.ColumnName = property.Name;
                                }
                                if (value == null)
                                {
                                    if (property.Name == "Code")
                                    {
                                        code = GetNextCode(TableName).PadLeft(10, '0');
                                        oGeneralData.SetProperty("Code", code);
                                        continue;
                                    }
                                    else
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
                                }

                                if (property.PropertyType != typeof(decimal) && property.PropertyType != typeof(Nullable<decimal>))
                                {
                                    oGeneralData.SetProperty(hubModel.ColumnName, value);
                                }
                                else
                                {
                                    oGeneralData.SetProperty(hubModel.ColumnName, Convert.ToDouble(value));
                                }
                                break;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception(String.Format("Erro ao setar propriedade {0}: {1}", property.Name, e));
                    }
                }

                switch (enumCrudOperation)
                {
                    case EnumCrudOperation.Create:
                        oGeneralService.Add(oGeneralData);
                        break;
                    case EnumCrudOperation.Update:
                        oGeneralService.Update(oGeneralData);
                        break;
                    default:
                        break;
                }







                return code.ToString();

            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (sCompany != null)
                {
                    Marshal.ReleaseComObject(sCompany);
                    sCompany = null;
                }

                if (oGeneralService != null)
                {
                    Marshal.ReleaseComObject(oGeneralService);
                    oGeneralService = null;
                }

                if (oGeneralData != null)
                {
                    Marshal.ReleaseComObject(oGeneralData);
                    oGeneralData = null;
                }
            }
        }

        /// <summary>
        /// Salva o model no BD de acordo com o tipo da operação
        /// </summary>
        /// <param name="enumCrudOperation">Operação - Create ou Update</param>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="model">Modelo</param>
        private string SaveNonObjectModel(EnumCrudOperation enumCrudOperation)
        {
            UserTable utbUser = SBOApp.Company.UserTables.Item(TableName.Replace("@", ""));
            try
            {
                if (UserTableType == BoUTBTableType.bott_NoObject)



                {
                    string code = String.Empty;
                    bool alreadyExists = false;
                    PropertyInfo propCode = Model.GetType().GetProperty("Code");
                    if (propCode != null)
                    {
                        if (propCode.GetValue(Model, null) != null)
                        {
                            code = propCode.GetValue(Model, null).ToString();
                            alreadyExists = utbUser.GetByKey(code);
                        }
                    }
                    if (!String.IsNullOrEmpty(code))
                    {
                        utbUser.Code = code;
                    }
                    else
                    {
                        utbUser.Code = CrudDAO.GetNextCode(TableName).PadLeft(10, '0');

                    }
                }









                PropertyInfo propName = Model.GetType().GetProperty("Name");
                if (propName != null && propName.GetValue(Model, null) != null)
                {
                    string name = propName.GetValue(Model, null).ToString();
                    if (name.Length > 100)
                    {
                        name = name.Substring(0, 100);
                    }
                    utbUser.Name = name;
                }
                else
                {
                    utbUser.Name = utbUser.Code;
                }

                HubModelAttribute hubModel;
                // Percorre as propriedades do Model
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    if (property.Name == "Code" || property.Name == "Name")
                    {
                        continue;
                    }

                    // Busca os Custom Attributes
                    foreach (Attribute attribute in property.GetCustomAttributes(true))
                    {
                        hubModel = attribute as HubModelAttribute;

                        if (hubModel != null)
                        {
                            // Se não for DataBaseField não seta nas properties
                            if (!hubModel.DataBaseFieldYN)
                            {
                                break;
                            }
                            if (String.IsNullOrEmpty(hubModel.ColumnName))
                                hubModel.ColumnName = property.Name;
                            object value = property.GetValue(Model, null);
                            if (value != null)
                            {
                                utbUser.UserFields.Fields.Item(hubModel.ColumnName).Value = value;
                            }
                        }
                        break;
                    }
                }

                int error = 0;
                switch (enumCrudOperation)
                {
                    case EnumCrudOperation.Create:
                        error = utbUser.Add();
                        break;
                    case EnumCrudOperation.Update:
                        error = utbUser.Update();
                        break;
                    case EnumCrudOperation.Delete:
                        error = utbUser.Remove();
                        break;
                    default:
                        break;
                }

                if (error != 0)
                {
                    throw new Exception(SBOApp.Company.GetLastErrorDescription());
                }

                return utbUser.Code;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(utbUser);
                utbUser = null;
                GC.Collect();
            }
        }

        #endregion CreateUpdateModel

        public string GetSqlCommand(Type modelType, string where, string orderBy, bool getValidValues)
        {
            Dictionary<string, Dictionary<string, string>> fieldValidValues = new Dictionary<string, Dictionary<string, string>>();
            if (getValidValues)
            {
                string sqlValidValues = @"SELECT CUFD.AliasID, UFD1.FldValue, UFD1.Descr FROM CUFD
                                    INNER JOIN UFD1
	                                    ON UFD1.TableID = CUFD.TableID
	                                    AND UFD1.FieldID = CUFD.FieldID
                                    WHERE CUFD.TableID = '{0}'";
                sqlValidValues = String.Format(sqlValidValues, TableName);

                Recordset rstValidValues = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                sqlValidValues = SBOApp.TranslateToHana(sqlValidValues);
                rstValidValues.DoQuery(sqlValidValues);
                string aliasId = null;
                while (!rstValidValues.EoF)
                {
                    if (String.IsNullOrEmpty(aliasId))
                    {
                        aliasId = rstValidValues.Fields.Item("AliasID").Value.ToString();
                    }

                    Dictionary<string, string> validValues = new Dictionary<string, string>();
                    while (!rstValidValues.EoF && aliasId == rstValidValues.Fields.Item("AliasID").Value.ToString())
                    {
                        validValues.Add(rstValidValues.Fields.Item("FldValue").Value.ToString(), rstValidValues.Fields.Item("Descr").Value.ToString());
                        rstValidValues.MoveNext();
                    }
                    fieldValidValues.Add(aliasId, validValues);
                    if (!rstValidValues.EoF)
                    {
                        aliasId = rstValidValues.Fields.Item("AliasID").Value.ToString();
                    }
                }
            }
            StringBuilder sql = new StringBuilder();
            sql.Append(" SELECT ");
            HubModelAttribute hubModel;
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;
                    // Se propriedade "ColumnName" estiver vazia, pega o nome da propriedade
                    if (String.IsNullOrEmpty(hubModel.Description))
                    {
                        hubModel.Description = property.Name;
                    }
                    if (String.IsNullOrEmpty(hubModel.ColumnName))
                    {
                        hubModel.ColumnName = property.Name;
                    }

                    if (!getValidValues || !fieldValidValues.ContainsKey(hubModel.ColumnName.Replace("U_", "")))
                    {
                        try
                        {
                            if (SBOApp.Company.DbServerType == (BoDataServerTypes)9)
                            {
                                sql.AppendFormat(", {0} ", hubModel.ColumnName, hubModel.Description);
                            }
                            else
                            {
                                sql.AppendFormat(", {0} AS '{1}' ", hubModel.ColumnName, hubModel.Description);
                            }
                        }
                        catch
                        {
                            sql.AppendFormat(", {0} AS '{1}' ", hubModel.ColumnName, hubModel.Description);
                        }
                        sql.AppendLine();
                    }
                    else
                    {
                        sql.AppendFormat(", CASE CAST({0} AS NVARCHAR) ", hubModel.ColumnName);
                        sql.AppendLine();
                        foreach (string strKey in fieldValidValues[hubModel.ColumnName.Replace("U_", "")].Keys)
                        {
                            sql.AppendFormat(" WHEN '{0}' THEN '{1}' ", strKey, fieldValidValues[hubModel.ColumnName.Replace("U_", "")][strKey]);
                            sql.AppendLine();
                        }
                        sql.AppendFormat(" END AS {0} ", hubModel.Description);
                        sql.AppendLine();
                    }
                }
            }
            sql.AppendFormat(" FROM [{0}]", TableName);
            sql.AppendLine();
            if (!String.IsNullOrEmpty(where))
            {
                sql.AppendFormat(" WHERE {0} ", where);
                sql.AppendLine();
            }
            if (!String.IsNullOrEmpty(orderBy))
            {
                sql.AppendFormat(" ORDER BY {0} ", orderBy);
                sql.AppendLine();
            }

            return sql.ToString().Replace("SELECT ,", "SELECT ");
        }

        #region RetrieveModel

        public T RetrieveModelByKey<T>(string key)
        {
            T model = Activator.CreateInstance<T>();
            HubModelAttribute hubModel;
            UserTable utbModel = SBOApp.Company.UserTables.Item(TableName.Replace("@", ""));

            if (utbModel.GetByKey(key))
            {
                foreach (PropertyInfo property in model.GetType().GetProperties())
                {
                    // Busca os Custom Attributes
                    foreach (Attribute attribute in property.GetCustomAttributes(true))
                    {
                        hubModel = attribute as HubModelAttribute;
                        // Se propriedade "ColumnName" estiver vazia, pega o nome da propriedade
                        if (String.IsNullOrEmpty(hubModel.ColumnName))
                            hubModel.ColumnName = property.Name;
                        if (hubModel != null)
                        {
                            if (hubModel.DataBaseFieldYN)
                            {
                                property.SetValue(model, utbModel.UserFields.Fields.Item(hubModel.ColumnName).Value, null);
                            }
                        }
                    }
                }
            }

            Marshal.ReleaseComObject(utbModel);
            utbModel = null;
            GC.Collect();

            return model;
        }
        /// <summary>
        /// Retorna Model preenchido de acordo com a condição WHERE
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição da consulta</param>
        /// <returns>Model</returns>
        public T RetrieveModel<T>(string where)
        {
            return this.RetrieveModel<T>(where, String.Empty);
        }

        /// <summary>
        /// Retorna Model preenchido de acordo com a condição WHERE
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição da consulta</param>
        /// <param name="orderBy">Ordenação</param>
        /// <returns>Model</returns>
        public T RetrieveModel<T>(string where, string orderBy)
        {
            List<T> modelList = this.RetrieveModelList<T>(where, orderBy);
            if (modelList.Count > 0)
                return modelList[0];
            else
                return Activator.CreateInstance<T>();
        }

        /// <summary>
        /// Retorna lista de Models de acordo com a condição WHERE
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição da consulta</param>
        /// <returns>ModelList</returns>
        public List<T> RetrieveModelList<T>(string where)
        {
            return this.RetrieveModelList<T>(String.Empty, String.Empty, where, String.Empty);
        }

        /// <summary>
        /// Retorna lista de Models de acordo com a condição WHERE
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição da consulta</param>
        /// <param name="orderBy">Ordenação</param>
        /// <returns>ModelList</returns>
        public List<T> RetrieveModelList<T>(string where, string orderBy)
        {
            return this.RetrieveModelList<T>(String.Empty, String.Empty, where, orderBy);
        }

        /// <summary>
        /// Retorna lista de Models de acordo com a condição WHERE
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="tableName">Nome da tabela</param>
        /// /// <param name="joinTable">Tabela </param>
        /// /// <param name="joinCondition">Nome da tabela</param>
        /// <param name="where">Condição da consulta</param>
        /// <param name="orderBy">Ordenação</param>
        /// <returns>ModelList</returns>
        public List<T> RetrieveModelList<T>(string joinTable, string joinCondition, string where, string orderBy)
        {
            StringBuilder sql = new StringBuilder();
            // Inicia o SELECT
            sql.Append(" SELECT ");

            Type modelType = typeof(T);
            HubModelAttribute hubModel;

            string fields = String.Empty;
            string fieldTableName;
            // Percorre as propriedades do Model para montar o SELECT
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;
                    if (hubModel == null)
                    {
                        continue;
                    }
                    // Se propriedade "ColumnName" estiver vazia, pega o nome da propriedade
                    if (String.IsNullOrEmpty(hubModel.ColumnName))
                        hubModel.ColumnName = property.Name;
                    if (hubModel != null)
                    {
                        // Se não for DataBaseField não adiciona no select
                        if (!hubModel.DataBaseFieldYN)
                        {
                            break;
                        }
                        fieldTableName = String.IsNullOrEmpty(hubModel.TableName) ? TableName : hubModel.TableName;
                        if (SBOApp.Company.DbServerType == (BoDataServerTypes)9)
                        {
                            fields += String.Format(", {1} ", fieldTableName, hubModel.ColumnName);
                        }
                        else
                        {
                            fields += String.Format(", [{0}].{1} AS {1} ", fieldTableName, hubModel.ColumnName);
                        }
                    }
                    break;
                }
            }

            if (String.IsNullOrEmpty(fields))
            {
                throw new Exception("Nenhuma propriedade do tipo hubModel encontrada no Model");
            }

            // Campos a serem retornados
            sql.Append(fields.Substring(1));

            // TABELA

            if (SBOApp.Company.DbServerType == (BoDataServerTypes)9)
            {
                sql.AppendFormat(" FROM [{0}] ", TableName);
            }
            else
            {
                sql.AppendFormat(" FROM [{0}] WITH(NOLOCK) ", TableName);
            }
            // INNER JOIN
            if (!String.IsNullOrEmpty(joinTable))
            {
                sql.AppendFormat(" INNER JOIN {0} ", joinTable);
                if (String.IsNullOrEmpty(joinCondition))
                {
                    joinCondition = " 1 = 1 ";
                }
                sql.AppendFormat(" ON {0} ", joinCondition);
            }

            // Condição WHERE
            if (!String.IsNullOrEmpty(where))
            {
                sql.AppendFormat(" WHERE {0} ", where);
            }

            // Condição ORDER BY
            if (!String.IsNullOrEmpty(orderBy))
            {
                sql.AppendFormat(" ORDER BY {0} ", orderBy);
            }

            return FillModelList<T>(sql.ToString());
        }
        #endregion RetrieveModel

        #region Delete
        /// <summary>
        /// Deleta registro
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição WHERE</param>
        public void DeleteModel(string tableName, string where)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            string sql = @"SELECT Code FROM [{0}] WHERE {1}";

            sql = SBOApp.TranslateToHana(String.Format(sql, tableName, where));

            rs.DoQuery(String.Format(sql, tableName, where));
            if (rs.RecordCount > 0)
            {
                CompanyService sCompany = SBOApp.Company.GetCompanyService();
                GeneralService oGeneralService = sCompany.GetGeneralService(tableName.Replace("@", ""));

                GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", rs.Fields.Item(0).Value.ToString());

                oGeneralService.Delete(oGeneralParams);
            }

            //Libera o objeto rs e chama o Garbage Collector
            Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();

        }

        public void DeleteModelByCode(string tableName, string code)
        {
            CompanyService sCompany = SBOApp.Company.GetCompanyService();
            GeneralService oGeneralService = sCompany.GetGeneralService(tableName.Replace("@", ""));

            GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("Code", code);

            oGeneralService.Delete(oGeneralParams);
        }
        #endregion Delete
        #endregion CRUD

        #region Util
        #region GetNextCode
        /// <summary>
        /// Retorna o próximo código
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <returns>Código</returns>
        public static string GetNextCode(string tableName)
        {
            return GetNextCode(tableName, "Code", String.Empty);
        }

        /// <summary>
        /// Retorna o próximo código
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="fieldName">Nome do campo</param>
        /// <returns>Código</returns>
        public static string GetNextCode(string tableName, string fieldName)
        {
            return GetNextCode(tableName, fieldName, String.Empty);
        }

        /// <summary>
        /// Retorna o próximo código
        /// </summary>
        /// <param name="fieldName">Nome do campo</param>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Where</param>
        /// <returns>Código</returns>
        public static string GetNextCode(string tableName, string fieldName, string where)
        {
            string sSql = String.Format(" SELECT ISNULL(MAX(CAST({0} AS NUMERIC(19, 6))), 0) + 1 FROM [{1}] ", fieldName, tableName);

            if (!String.IsNullOrEmpty(where))
            {
                sSql += String.Format(" WHERE {0} ", where);
            }

            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            sSql = SBOApp.TranslateToHana(sSql);
            rs.DoQuery(sSql);
            string code = rs.Fields.Item(0).Value.ToString();

            //Libera o objeto rs e chama o Garbage Collector
            Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();

            return code;
        }
        #endregion GetNextCode

        #region GetColumnValue
        /// <summary>
        /// Retorna valor da coluna de acordo com o select
        /// </summary>
        public static T GetColumnValue<T>(string sql)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                sql = SBOApp.TranslateToHana(sql);
                rs.DoQuery(sql);
                T value = (T)rs.Fields.Item(0).Value;
                return value;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(rs);
                rs = null;
            }
        }
        #endregion

        public static List<string> FillStringList(string sql)
        {
            Recordset rst = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            List<string> list = new List<string>();
            try
            {
                rst.DoQuery(sql);

                while (!rst.EoF)
                {
                    list.Add(rst.Fields.Item(0).Value.ToString());
                    rst.MoveNext();
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (rst != null)
                {
                    //Libera o objeto rs e chama o Garbage Collector
                    Marshal.ReleaseComObject(rst);
                    rst = null;
                }
            }

            return list;
        }

        public int GetRowCount(string sql)
        {
            Recordset rst = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rst.DoQuery(sql);
                return rst.RecordCount;
            }
            finally
            {
                Marshal.ReleaseComObject(rst);
                rst = null;
            }
        }

        #region FillModel
        /// <summary>
        /// Preenche model através do SQL
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public T FillModel<T>(string sql)
        {
            List<T> modelList = this.FillModelList<T>(sql);
            if (modelList.Count > 0)
                return modelList[0];
            else
                return Activator.CreateInstance<T>();
        }

        /// <summary>
        /// Preenche a lista de model através do SQL
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public List<T> FillModelList<T>(string sql)
        {
            List<T> modelList = new List<T>();
            T model;
            HubModelAttribute hubModel;
            Recordset rs = null;
            try
            {
                // Lê os dados em um Recordset
                rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    sql = SBOApp.TranslateToHana(sql);
                }
                catch { }
                rs.DoQuery(sql);

                // Lê os dados e insere no model
                if (rs.RecordCount > 0)
                {
                    while (!rs.EoF)
                    {
                        // Cria nova instância do model
                        model = Activator.CreateInstance<T>();
                        // Seta os valores no model
                        foreach (PropertyInfo property in model.GetType().GetProperties())
                        {
                            try
                            {
                                // Busca os Custom Attributes
                                foreach (Attribute attribute in property.GetCustomAttributes(true))
                                {
                                    hubModel = attribute as HubModelAttribute;
                                    if (hubModel != null)
                                    {
                                        // Se propriedade "ColumnName" estiver vazia, pega o nome da propriedade
                                        if (String.IsNullOrEmpty(hubModel.ColumnName))
                                            hubModel.ColumnName = property.Name;
                                        // Se não for DataBaseField não seta nas properties
                                        if (!hubModel.DataBaseFieldYN && !hubModel.FillOnSelect)
                                        {
                                            break;
                                        }
                                        if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(Nullable<decimal>))
                                        {
                                            property.SetValue(model, Convert.ToDecimal(rs.Fields.Item(hubModel.ColumnName).Value), null);
                                        }
                                        else if (property.PropertyType == typeof(short) || property.PropertyType == typeof(Nullable<short>))
                                        {
                                            property.SetValue(model, Convert.ToInt16(rs.Fields.Item(hubModel.ColumnName).Value), null);
                                        }
                                        else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(Nullable<bool>))
                                        {
                                            property.SetValue(model, Convert.ToBoolean(rs.Fields.Item(hubModel.ColumnName).Value), null);
                                        }
                                        else
                                        {
                                            try
                                            {
                                                property.SetValue(model, rs.Fields.Item(hubModel.ColumnName).Value, null);
                                            }
                                            catch (Exception e)
                                            {
                                                if (property.PropertyType == typeof(DateTime))
                                                {
                                                    try
                                                    {
                                                        string dateStr = rs.Fields.Item(hubModel.ColumnName).Value.ToString();
                                                        DateTime date;

                                                        if (DateTime.TryParseExact(dateStr, "dd-MM-yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                                                        {
                                                            property.SetValue(model, date, null);
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {
                                                            string hour = rs.Fields.Item(hubModel.ColumnName).Value.ToString().PadLeft(4, '0');
                                                            DateTime date;

                                                            if (DateTime.TryParseExact(hour, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                                                            {
                                                                property.SetValue(model, date, null);
                                                            }
                                                        }
                                                        catch
                                                        {
                                                            throw e;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    throw e;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                throw new Exception(String.Format("Erro ao setar propriedade {0}: {1}", property.Name, e));
                            }
                        }

                        // Adiciona na lista
                        modelList.Add(model);
                        rs.MoveNext();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (rs != null)
                {
                    //Libera o objeto rs e chama o Garbage Collector
                    Marshal.ReleaseComObject(rs);
                    rs = null;
                }
            }

            return modelList;
        }

        /// <summary>
        /// Preenche model através do SQL
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public T FillModelFromSql<T>(string sql)
        {
            List<T> modelList = this.FillModelListFromSql<T>(sql);
            if (modelList.Count > 0)
                return modelList[0];
            else
                return Activator.CreateInstance<T>();
        }

        /// <summary>
        /// Preenche as propriedades do model de acordo com as colunas do SELECT
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public List<T> FillModelListFromSql<T>(string sql)
        {
            List<T> modelList = new List<T>();
            T model;
            Recordset rs = null;
            try
            {
                // Lê os dados em um Recordset
                rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                sql = SBOApp.TranslateToHana(sql);
                rs.DoQuery(sql);

                // Lê os dados e insere no model
                if (rs.RecordCount > 0)
                {
                    while (!rs.EoF)
                    {
                        // Cria nova instância do model
                        model = Activator.CreateInstance<T>();
                        Type modelType = model.GetType();
                        // Seta os valores no model
                        for (int i = 0; i < rs.Fields.Count; i++)
                        {
                            if (rs.Fields.Item(i).IsNull() == BoYesNoEnum.tYES && rs.Fields.Item(i).Value.ToString() == "")
                            {
                                continue;
                            }

                            PropertyInfo property = modelType.GetProperty(rs.Fields.Item(i).Name);
                            try
                            {
                                if (property == null)
                                {
                                    throw new Exception($"Campo {rs.Fields.Item(i).Name} não encontrado no model {modelType.Name}");
                                }

                                if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(Nullable<decimal>))
                                {
                                    property.SetValue(model, Convert.ToDecimal(rs.Fields.Item(i).Value), null);
                                }
                                else if (property.PropertyType == typeof(short) || property.PropertyType == typeof(Nullable<short>))
                                {
                                    property.SetValue(model, Convert.ToInt16(rs.Fields.Item(i).Value), null);
                                }
                                else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(Nullable<bool>))
                                {
                                    property.SetValue(model, Convert.ToBoolean(rs.Fields.Item(i).Value), null);
                                }
                                else if (property.PropertyType == typeof(byte[]) || property.PropertyType == typeof(Nullable<byte>))
                                {
                                    string hex = rs.Fields.Item(i).Value.ToString();
                                    hex = hex.SafeSubstring(2, hex.Length);
                                    if (hex.Length % 2 == 1) hex = "0" + hex;
                                    byte[] bytes = System.Runtime.Remoting.Metadata.W3cXsd2001.SoapHexBinary.Parse(hex).Value;
                                    property.SetValue(model, bytes, null);
                                }
                                else
                                {
                                    try
                                    {
                                        property.SetValue(model, rs.Fields.Item(i).Value, null);
                                    }
                                    catch (Exception e)
                                    {
                                        if (property.PropertyType == typeof(DateTime))
                                        {
                                            try
                                            {
                                                string hour = rs.Fields.Item(i).Value.ToString().PadLeft(4, '0');
                                                DateTime date;

                                                if (DateTime.TryParseExact(hour, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                                                {
                                                    property.SetValue(model, date, null);
                                                }
                                            }
                                            catch
                                            {
                                                throw e;
                                            }
                                        }
                                        else
                                        {
                                            throw e;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"{modelType.Name} - {rs.Fields.Item(i).Name}: {ex.Message}");
                            }
                        }

                        // Adiciona na lista
                        modelList.Add(model);
                        rs.MoveNext();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (rs != null)
                {
                    //Libera o objeto rs e chama o Garbage Collector
                    Marshal.ReleaseComObject(rs);
                    rs = null;
                }
            }

            return modelList;
        }
        #endregion FillModel

        #region Exists
        /// <summary>
        /// Verifica se registro existe
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição WHERE</param>
        /// <returns>Código do registro</returns>
        public string Exists(string where)
        {
            return this.Exists("Code", where);
        }

        /// <summary>
        /// Verifica se registro existe
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="returnColumn">Coluna a ser retornada</param>
        /// <param name="where">Condição WHERE</param>
        /// <returns>Código do registro</returns>
        public string Exists(string returnColumn, string where)
        {
            string sql = String.Format("SELECT TOP 1 {0} FROM [{1}] WITH(NOLOCK) ", returnColumn, TableName);

            if (!String.IsNullOrEmpty(where))
            {
                sql += String.Format(" WHERE {0} ", where);
            }

            // Lê os dados em um Recordset
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sql = SBOApp.TranslateToHana(sql);
            rs.DoQuery(sql);
            if (rs.RecordCount > 0)
            {
                returnColumn = rs.Fields.Item(0).Value.ToString();

                //Libera o objeto rs e chama o Garbage Collector
                Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();

                return returnColumn;
            }
            else
            {
                return null;
            }
        }
        #endregion Exists

        public static void ExecuteNonQuery(string sql)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sql = SBOApp.TranslateToHana(sql);
            rs.DoQuery(sql);

            Marshal.ReleaseComObject(rs);
            rs = null;
        }

        public static object ExecuteScalar(string sql)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sql = SBOApp.TranslateToHana(sql);
            rs.DoQuery(sql);
            object ret;

            if (rs.RecordCount > 0)
            {
                ret = rs.Fields.Item(0).Value;

                Marshal.ReleaseComObject(rs);
                rs = null;

                return ret;
            }
            else
            {
                return null;
            }
        }
        #endregion Util
    }
}
