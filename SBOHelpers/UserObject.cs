using SBO.Hub.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace SBO.Hub.Helpers
{
    /// <summary>
    /// ATENÇÂO - Cuidado ao efetuar alterações:
    /// Não é possível iniciar um Recordset (ou algum outro objeto) junto com os objetos de criação de campos/tabelas
    /// </summary>
    public class UserObject
    {
        public List<string> LogList { get; set; }
        int CodErro;
        string MsgErro;
        public bool? Confirmed;
        public bool AskConfirmation = false;
        //private GenericModel FindColumns;

        public UserObject(bool askConfirmation = false)
        {
            LogList = new List<string>();
            AskConfirmation = askConfirmation;
        }

        public void CreateUserTable(string UserTableName, string UserTableDesc, BoUTBTableType UserTableType)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserTableName = UserTableName.Replace("@", "");
            UserTablesMD oUserTableMD = (UserTablesMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserTables);

            try
            {
                bool update = oUserTableMD.GetByKey(UserTableName);
                if (update)
                {
                    return;
                }

                oUserTableMD.TableName = UserTableName;
                oUserTableMD.TableDescription = UserTableDesc;
                oUserTableMD.TableType = UserTableType;

                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                    else
                    {
                        AskConfirmation = false;
                    }
                }
                CodErro = oUserTableMD.Add();
                this.ValidateAction();
            }
            catch (Exception ex)
            {
                LogList.Add("Erro geral ao criar tabela: " + ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oUserTableMD);
                oUserTableMD = null;
                GC.Collect();
            }
        }

        public void RemoveUserTable(string UserTableName)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserTablesMD oUserTableMD = (UserTablesMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserTables);

            // Remove a arroba do usertable Name
            UserTableName = UserTableName.Replace("@", "");

            if (oUserTableMD.GetByKey(UserTableName))
            {
                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                    else
                    {
                        AskConfirmation = false;
                    }
                }
                CodErro = oUserTableMD.Remove();
                this.ValidateAction();
            }
            else
            {
                CodErro = 0;
                MsgErro = "";
            }
            Marshal.ReleaseComObject(oUserTableMD);
            oUserTableMD = null;
            GC.Collect();
        }

        public void InsertUserField(string TableName, string FieldName, string FieldDescription, BoFieldTypes oType, BoFldSubTypes oSubType, int FieldSize, bool MandatoryYN = false, string DefaultValue = "", string linkedTable = "")
        {
            if (Confirmed == false)
            {
                return;
            }
            if (FieldDescription.Length > 30)
            {
                FieldDescription = FieldDescription.Substring(0, 30);
            }

            string Sql = " SELECT FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
            Sql = String.Format(Sql, TableName, FieldName);
            string FieldId = QueryForValue(Sql);

            if (FieldId != null)
            {
                return;
            }
            if (AskConfirmation)
            {
                if (Confirmation())
                {
                    return;
                }
                else
                {
                    AskConfirmation = false;
                }
            }

            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));
            oUserFieldsMD.TableName = TableName;
            oUserFieldsMD.Name = FieldName;
            oUserFieldsMD.Description = FieldDescription;
            oUserFieldsMD.Type = oType;
            oUserFieldsMD.SubType = oSubType;
            oUserFieldsMD.Mandatory = GetSapBoolean(MandatoryYN);

            if (!String.IsNullOrEmpty(DefaultValue))
            {
                oUserFieldsMD.DefaultValue = DefaultValue;
            }
            if (!String.IsNullOrEmpty(linkedTable))
            {
                oUserFieldsMD.LinkedTable = linkedTable;
            }

            if (FieldSize > 0)
                oUserFieldsMD.EditSize = FieldSize;
            CodErro = oUserFieldsMD.Add();
            this.ValidateAction();

            Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
        }


        public void UpsertUserField(string TableName, string FieldName, string FieldDescription, BoFieldTypes oType, BoFldSubTypes oSubType, int FieldSize, bool MandatoryYN = false, string DefaultValue = "")
        {
            if (Confirmed == false)
            {
                return;
            }
            if (FieldDescription.Length > 30)
            {
                FieldDescription = FieldDescription.Substring(0, 30);
            }

            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));
            bool bUpdate;

            string Sql = " SELECT FieldId FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
            Sql = String.Format(Sql, TableName, FieldName);
            string FieldId = QueryForValue(Sql);

            if (FieldId != null)
            {
                bUpdate = oUserFieldsMD.GetByKey(TableName, Convert.ToInt32(FieldId));
            }
            else
                bUpdate = false;

            oUserFieldsMD.TableName = TableName;
            oUserFieldsMD.Name = FieldName;
            oUserFieldsMD.Description = FieldDescription;
            oUserFieldsMD.Type = oType;
            oUserFieldsMD.SubType = oSubType;
            oUserFieldsMD.Mandatory = GetSapBoolean(MandatoryYN);
            if (!String.IsNullOrEmpty(DefaultValue))
            {
                oUserFieldsMD.DefaultValue = DefaultValue;
            }

            if (FieldSize > 0)
                oUserFieldsMD.EditSize = FieldSize;

            if (bUpdate)
                //CodErro = oUserFieldsMD.Update();
                CodErro = 0;
            else
                CodErro = oUserFieldsMD.Add();
            this.ValidateAction();

            Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
        }

        public void RemoveUserField(string TableName, string FieldName)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));

            string Sql = " SELECT FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
            Sql = String.Format(Sql, TableName, FieldName);

            string FieldId = QueryForValue(Sql);

            if (FieldId != null)
            {
                if (oUserFieldsMD.GetByKey(TableName, Convert.ToInt32(FieldId)))
                {
                    if (AskConfirmation)
                    {
                        if (Confirmation())
                        {
                            return;
                        }
                        else
                        {
                            AskConfirmation = false;
                        }
                    }

                    CodErro = oUserFieldsMD.Remove();
                    this.ValidateAction();
                }
            }
            else
            {
                MsgErro = "";
                CodErro = 0;
                LogList.Add(" Tabela/Campo não encontrado ");
            }

            Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
        }

        public void AddValidValueToUserField(string TableName, string FieldName, Dictionary<string, string> values, string defaultValue = "")
        {
            if (Confirmed == false)
            {
                return;
            }
            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));
            try
            {
                string sql = @" SELECT FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
                sql = String.Format(sql, TableName, FieldName.Replace("U_", ""));
                string FieldId = QueryForValue(sql);

                if (!oUserFieldsMD.GetByKey(TableName, Convert.ToInt32(FieldId)))
                {
                    LogList.Add($"AddValidValueToUserField - Campo {FieldName} não encontrado na tabela {TableName}");
                }

                bool update = false;

                foreach (var item in values)
                {
                    sql = @" SELECT UFD1.IndexID FROM CUFD
                            INNER JOIN UFD1 
                                ON CUFD.TableID = UFD1.TableID 
                                AND CUFD.FieldID = UFD1.FieldID
                         WHERE CUFD.TableID = '{0}' 
                         AND CUFD.AliasID = '{1}' 
                         AND UFD1.FldValue = '{2}'";
                    sql = String.Format(sql, TableName, FieldName.Replace("U_", ""), item.Key);

                    string IndexId = QueryForValue(sql);

                    if (IndexId == null)
                    {
                        update = true;
                        if (!String.IsNullOrEmpty(oUserFieldsMD.ValidValues.Value))
                        {
                            oUserFieldsMD.ValidValues.Add();
                        }
                    }
                    else
                    {
                        continue;
                    }

                    oUserFieldsMD.ValidValues.Value = item.Key;
                    oUserFieldsMD.ValidValues.Description = item.Value;

                    if (item.Key == defaultValue)
                    {
                        oUserFieldsMD.DefaultValue = defaultValue;
                    }
                }

                if (update)
                {
                    if (AskConfirmation)
                    {
                        if (Confirmation())
                        {
                            return;
                        }
                        else
                        {
                            AskConfirmation = false;
                        }
                    }

                    CodErro = oUserFieldsMD.Update();
                    this.ValidateAction();
                }
            }
            catch (Exception ex)
            {
                LogList.Add($"Erro geral ao inserir valor válido: {ex.Message}");
            }
            finally
            {
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }
        }

        public void AddValidValueToUserField(string TableName, string FieldName, string Value, string Description)
        {
            // se não foi passado o parâmetro de "É Valor Padrão" trata como não
            // chamando a função que realmente insere o valor como "false" a variável IsDefault
            AddValidValueToUserField(TableName, FieldName, Value, Description, false);
        }

        public void AddValidValueToUserField(string TableName, string FieldName, string Value, string Description, bool IsDefault)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));
            try
            {
                string sql = @" SELECT UFD1.IndexID FROM CUFD
                            INNER JOIN UFD1 
                                ON CUFD.TableID = UFD1.TableID 
                                AND CUFD.FieldID = UFD1.FieldID
                         WHERE CUFD.TableID = '{0}' 
                         AND CUFD.AliasID = '{1}' 
                         AND UFD1.FldValue = '{2}'";
                sql = String.Format(sql, TableName, FieldName.Replace("U_", ""), Value);

                string IndexId = QueryForValue(sql);

                if (IndexId != null)
                {
                    return;
                }

                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                    else
                    {
                        AskConfirmation = false;
                    }
                }

                sql = @" SELECT FieldID FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
                sql = String.Format(sql, TableName, FieldName.Replace("U_", ""));
                string FieldId = QueryForValue(sql);

                if (!oUserFieldsMD.GetByKey(TableName, Convert.ToInt32(FieldId)))
                {
                    throw new Exception("Campo não encontrado!");
                }

                if (!String.IsNullOrEmpty(oUserFieldsMD.ValidValues.Value))
                {
                    oUserFieldsMD.ValidValues.Add();
                }

                oUserFieldsMD.ValidValues.Value = Value;
                oUserFieldsMD.ValidValues.Description = Description;

                if (IsDefault)
                    oUserFieldsMD.DefaultValue = Value;

                CodErro = oUserFieldsMD.Update();

                this.ValidateAction();
            }
            catch (Exception ex)
            {
                LogList.Add($"Erro geral ao inserir valor válido: {ex.Message}");
            }
            finally
            {
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }
        }

        public void AddValidValueFromTable(string tableName, string fieldName, string baseTable, string valueField = "Code", string descriptionField = "Name")
        {
            if (Confirmed == false)
            {
                return;
            }
            UserFieldsMD oUserFieldsMD = ((UserFieldsMD)(SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserFields)));
            try
            {
                string sql = @" SELECT FieldId FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
                sql = String.Format(sql, tableName, fieldName.Replace("U_", ""));
                string fieldId = QueryForValue(sql);
                List<ValidValueModel> validValuesList = this.GetValidValues(tableName, fieldId, baseTable, valueField, descriptionField);
                bool update = false;
                if (validValuesList.Count > 0)
                {
                    oUserFieldsMD.GetByKey(tableName, Convert.ToInt32(fieldId));
                    foreach (var item in validValuesList)
                    {
                        if (!String.IsNullOrEmpty(oUserFieldsMD.ValidValues.Value))
                        {
                            oUserFieldsMD.ValidValues.Add();
                        }
                        if (oUserFieldsMD.ValidValues.Value != item.Value || oUserFieldsMD.ValidValues.Description != item.Description)
                        {
                            update = true;
                        }
                        oUserFieldsMD.ValidValues.Value = item.Value;
                        oUserFieldsMD.ValidValues.Description = item.Description;
                    }

                    if (update)
                    {
                        if (AskConfirmation)
                        {
                            if (Confirmation())
                            {
                                return;
                            }
                            else
                            {
                                AskConfirmation = false;
                            }
                        }
                        CodErro = oUserFieldsMD.Update();
                        this.ValidateAction();
                    }
                }
            }
            catch (Exception ex)
            {
                LogList.Add($"Erro geral ao inserir valor válido: {ex.Message}");
            }
            finally
            {
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }
        }

        public List<ValidValueModel> GetValidValues(string tableName, string fieldId, string baseTable, string valueField = "Code", string descriptionField = "Name")
        {
            Recordset oRecordset = (Recordset)(SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset));
            List<ValidValueModel> list = new List<ValidValueModel>();
            try
            {
                string sql = $@" SELECT {valueField}, {descriptionField} FROM [{baseTable}]
                                 WHERE NOT EXISTS
                                 (
                                    SELECT TOP 1 1 FROM UFD1
                                    WHERE TableID = '{tableName}'
                                        AND FieldID = {fieldId}
                                        AND FldValue = {valueField}
                                 ) ";
                sql = SBOApp.TranslateToHana(sql);
                oRecordset.DoQuery(sql);

                while (!oRecordset.EoF)
                {
                    ValidValueModel model = new ValidValueModel();
                    model.Value = oRecordset.Fields.Item(valueField).Value.ToString();
                    model.Description = oRecordset.Fields.Item(descriptionField).Value.ToString();
                    list.Add(model);
                    oRecordset.MoveNext();
                }
            }
            catch
            {

            }
            finally
            {
                Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
            }
            return list;
        }

        public void CreateUserObject(string ObjectName, string ObjectDesc, string TableName, BoUDOObjType ObjectType, bool CanLog = false, bool CanDelete = false, bool CanFind = true)
        {
            this.CreateUserObject(ObjectName, ObjectDesc, TableName, ObjectType, CanLog, CanDelete, CanFind, false, false, false, false, 0, 0, null);
        }

        public void CreateUserObject(string ObjectName, string ObjectDesc, string TableName, BoUDOObjType ObjectType, bool CanLog, bool CanDelete, bool CanFind, bool CanYearTransfer, bool CanCancel, bool CanClose, bool CanCreateDefaultForm, int FatherMenuId, int menuPosition, string srfFile = "")
        {
            if (Confirmed == false)
            {
                return;
            }
            // se não preenchido um table name separado, usa o mesmo do objeto
            if (String.IsNullOrEmpty(TableName))
                TableName = ObjectName;

            UserObjectsMD UserObjectsMD = (UserObjectsMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            try
            {
                // Remove a arroba do usertable Name
                TableName = TableName.Replace("@", "");

                if (UserObjectsMD.GetByKey(ObjectName))
                {
                    return;
                }

                UserObjectsMD.Code = ObjectName;
                UserObjectsMD.Name = ObjectDesc;
                UserObjectsMD.ObjectType = ObjectType;
                UserObjectsMD.TableName = TableName;

                //UserObjectsMD.CanArchive = GetSapBoolean(CanArchive);
                UserObjectsMD.CanCancel = GetSapBoolean(CanCancel);
                UserObjectsMD.CanClose = GetSapBoolean(CanClose);
                UserObjectsMD.CanCreateDefaultForm = GetSapBoolean(CanCreateDefaultForm);
                UserObjectsMD.CanDelete = GetSapBoolean(CanDelete);
                UserObjectsMD.CanFind = GetSapBoolean(CanFind);
                UserObjectsMD.CanLog = GetSapBoolean(CanLog);
                UserObjectsMD.CanYearTransfer = GetSapBoolean(CanYearTransfer);

                if (CanFind)
                {
                    UserObjectsMD.FindColumns.ColumnAlias = "Code";
                    UserObjectsMD.FindColumns.ColumnDescription = "Código";
                    UserObjectsMD.FindColumns.Add();
                }

                if (CanCreateDefaultForm)
                {
                    UserObjectsMD.CanCreateDefaultForm = BoYesNoEnum.tYES;
                    UserObjectsMD.CanCancel = GetSapBoolean(CanCancel);
                    UserObjectsMD.CanClose = GetSapBoolean(CanClose);
                    UserObjectsMD.CanDelete = GetSapBoolean(CanDelete);
                    UserObjectsMD.CanFind = GetSapBoolean(CanFind);
                    UserObjectsMD.ExtensionName = "";
                    UserObjectsMD.OverwriteDllfile = BoYesNoEnum.tYES;
                    UserObjectsMD.ManageSeries = BoYesNoEnum.tYES;
                    UserObjectsMD.UseUniqueFormType = BoYesNoEnum.tYES;
                    UserObjectsMD.EnableEnhancedForm = BoYesNoEnum.tNO;
                    UserObjectsMD.RebuildEnhancedForm = BoYesNoEnum.tNO;
                    UserObjectsMD.FormSRF = srfFile;

                    UserObjectsMD.FormColumns.FormColumnAlias = "Code";
                    UserObjectsMD.FormColumns.FormColumnDescription = "Código";
                    UserObjectsMD.FormColumns.Add();

                    UserObjectsMD.FatherMenuID = FatherMenuId;
                    UserObjectsMD.Position = menuPosition;
                    UserObjectsMD.MenuItem = BoYesNoEnum.tYES;
                    UserObjectsMD.MenuUID = ObjectName;
                    UserObjectsMD.MenuCaption = ObjectDesc;
                }

                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                    else
                    {
                        AskConfirmation = false;
                    }
                }
                CodErro = UserObjectsMD.Add();
                this.ValidateAction();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(UserObjectsMD);
                UserObjectsMD = null;
                GC.Collect();
            }
        }

        public void RemoveUserObject(string ObjectName)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserObjectsMD UserObjectsMD = (UserObjectsMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            if (UserObjectsMD.GetByKey(ObjectName))
            {
                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                    else
                    {
                        AskConfirmation = false;
                    }
                }

                CodErro = UserObjectsMD.Remove();
                this.ValidateAction();
            }
            Marshal.ReleaseComObject(UserObjectsMD);
            UserObjectsMD = null;
            GC.Collect();
        }

        public void AddChildTableToUserObject(string ObjectName, string ChildTableName)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserObjectsMD UserObjectsMD = (UserObjectsMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            // Remove a arroba do usertable Name
            ChildTableName = ChildTableName.Replace("@", "");

            bool bUpdate = UserObjectsMD.GetByKey(ObjectName);

            bool JaAdicionada = false;
            for (int i = 0; i < UserObjectsMD.ChildTables.Count; i++)
            {
                UserObjectsMD.ChildTables.SetCurrentLine(i);
                if (ChildTableName == UserObjectsMD.ChildTables.TableName)
                {
                    JaAdicionada = true;
                    break;
                }
            }

            if (!JaAdicionada)
            {
                if (AskConfirmation)
                {
                    if (Confirmation())
                    {
                        return;
                    }
                }

                UserObjectsMD.ChildTables.Add();
                UserObjectsMD.ChildTables.TableName = ChildTableName;

                CodErro = UserObjectsMD.Update();
                this.ValidateAction();
            }

            Marshal.ReleaseComObject(UserObjectsMD);
            UserObjectsMD = null;
            GC.Collect();
        }

        public static string QueryForValue(string Sql)
        {
            Recordset oRecordset = (Recordset)(SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset));
            string Retorno = null;
            try
            {
                Sql = SBOApp.TranslateToHana(Sql);
                oRecordset.DoQuery(Sql);

                // Executa e, caso exista ao menos um registro, devolve o mesmo.
                // retorna sempre o primeiro campo da consulta (SEMPRE)
                if (!oRecordset.EoF)
                {
                    Retorno = oRecordset.Fields.Item(0).Value.ToString();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();

            }

            return Retorno;
        }

        public static bool FieldExists(string tableName, string fieldName)
        {
            string sql = @" SELECT TOP 1 1 FROM CUFD WHERE TableID = '{0}' AND AliasID = '{1}' ";
            sql = String.Format(sql, tableName, fieldName.Replace("U_", ""));
            Recordset rst = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            sql = SBOApp.TranslateToHana(sql);

            rst.DoQuery(sql);
            bool exists = false;
            if (rst.RecordCount > 0)
            {
                exists = true;
            }

            Marshal.ReleaseComObject(rst);
            rst = null;
            GC.Collect();

            return exists;
        }

        public void CreateUserKey(string KeyName, string TableName, string Fields, bool isUnique)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserKeysMD oUserKeysMD = (UserKeysMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserKeys);

            oUserKeysMD.TableName = TableName;
            oUserKeysMD.KeyName = KeyName;

            string[] arrAux = Fields.Split(Convert.ToChar(","));
            for (int i = 0; i < arrAux.Length; i++)
            {
                if (i > 0)
                    oUserKeysMD.Elements.Add();

                oUserKeysMD.Elements.ColumnAlias = arrAux[i].Trim();
            }

            oUserKeysMD.Unique = GetSapBoolean(isUnique);

            CodErro = oUserKeysMD.Add();
            this.ValidateAction();

            Marshal.ReleaseComObject(oUserKeysMD);
            oUserKeysMD = null;
        }

        public void ValidateAction()
        {
            if (CodErro != 0)
            {
                SBOApp.Company.GetLastError(out CodErro, out MsgErro);
                LogList.Add($"FALHA ({MsgErro})");
            }
            else
            {
                MsgErro = "";
            }
        }

        public void MakeFieldsSearchable(string tableName)
        {
            if (Confirmed == false)
            {
                return;
            }
            UserObjectsMD userObjectsMD = (UserObjectsMD)SBOApp.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            Dictionary<string, string> fields = this.GetTableFields(tableName);

            tableName = tableName.Replace("@", "");
            if (userObjectsMD.GetByKey(tableName))
            {
                userObjectsMD.CanFind = BoYesNoEnum.tYES;
                bool hasNewColumn = false;
                foreach (var item in fields)
                {
                    bool found = false;
                    for (int i = 0; i < userObjectsMD.FindColumns.Count; i++)
                    {
                        userObjectsMD.FindColumns.SetCurrentLine(i);
                        if (userObjectsMD.FindColumns.ColumnAlias == item.Key)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        hasNewColumn = true;
                        userObjectsMD.FindColumns.ColumnAlias = item.Key;
                        userObjectsMD.FindColumns.ColumnDescription = item.Value;
                        userObjectsMD.FindColumns.Add();
                    }
                }

                if (hasNewColumn)
                {
                    if (AskConfirmation)
                    {
                        if (Confirmation())
                        {
                            return;
                        }
                    }

                    CodErro = userObjectsMD.Update();
                }

                this.ValidateAction();
            }
            Marshal.ReleaseComObject(userObjectsMD);
            userObjectsMD = null;

        }

        public void ShowLog()
        {
            if (LogList.Count > 0)
            {
                SBOApp.Application.MessageBox("Ocorreram erros na criação de campos de usuário. Verifique o log de mensagens do sistema");
                foreach (var log in LogList)
                {
                    SBOApp.Application.SetStatusBarMessage(log);
                }
            }
        }

        public Dictionary<string, string> GetTableFields(string tableName)
        {
            Recordset rs = (Recordset)(SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset));
            string sql = "SELECT * FROM CUFD WHERE TableID = '{0}'";

            Dictionary<string, string> fields = new Dictionary<string, string>();
            sql = SBOApp.TranslateToHana(String.Format(sql, tableName));
            rs.DoQuery(sql);
            while (!rs.EoF)
            {
                fields.Add("U_" + rs.Fields.Item("AliasID").Value.ToString(), rs.Fields.Item("Descr").Value.ToString());
                rs.MoveNext();
            }

            Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();
            return fields;
        }

        public static BoYesNoEnum GetSapBoolean(bool Variavel)
        {
            if (Variavel)
                return BoYesNoEnum.tYES;
            else
                return BoYesNoEnum.tNO;

        }

        private bool Confirmation()
        {
            Confirmed = SBOApp.Application.MessageBox("A aplicação irá adicionar campos de usuário no banco de dados. Deseja continuar?", 2, "Sim", "Não") == 1;
            return Confirmed.Value;
        }
    }
}
