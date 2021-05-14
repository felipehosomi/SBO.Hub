using SAPbobsCOM;
using SBO.Hub.Attributes;
using SBO.Hub.Enums;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace SBO.Hub.DAO 
{
    public class CrudDAO
    {
        public string TableName { get; set; }
        public object Model { get; set; }
        public BoUTBTableType UserTableType { get; set; }

        public CrudDAO()
        {
            
        }

        public CrudDAO(string tableName)
        {
            TableName = tableName;
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
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.UserTableType = UserTableType;
                DAO.TableName = TableName;
                DAO.Model = Model;
                return DAO.CreateModel();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.CreateModel();
                return String.Empty;
            }
        }

        /// <summary>
        /// Atualiza dados no banco
        /// </summary>
        /// <param name="model">Objeto do tipo Model</param>
        /// <param name="tableName">Nome da tabela</param>
        public void UpdateModel()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.UserTableType = UserTableType;
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.UpdateModel();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.UpdateModel();
            }
        }

        /// <summary>
        /// Atualiza dados no banco
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição WHRE</param>
        /// <param name="model">Model com os dados a serem atualizados</param>
        public void UpdateModel(string where)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.UserTableType = UserTableType;
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.UpdateModel(where);
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        #endregion CreateUpdateModel

        #region RetrieveModel
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
        public List<T> RetrieveModelList<T>(string where = "")
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
                        if (HubApp.DatabaseType == DatabaseTypeEnum.HANA)
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
                throw new Exception("Nenhuma propriedade do tipo ModelDAO encontrada no Model");
            }

            // Campos a serem retornados
            sql.Append(fields.Substring(1));

            // TABELA
            if (HubApp.DatabaseType == DatabaseTypeEnum.HANA)
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
        /// Remove registro
        /// </summary>
        public void DeleteModel()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.UserTableType = UserTableType;
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.DeleteModel();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                DAO.Model = Model;
                DAO.DeleteModel();
            }
        }

        /// <summary>
        /// Deleta registro
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <param name="where">Condição WHERE</param>
        public void DeleteModel(string tableName, string where)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.DeleteModel(tableName, where);
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        public void DeleteModelByCode(string tableName, string code)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.DeleteModelByCode(tableName, code);
            }
            else
            {
                throw new NotImplementedException();
            }
        }
        #endregion Delete
        #endregion CRUD

        #region Util
        #region GetSqlCommand
        /// <summary>
        /// Retorna command SQL montado de acordo com o model
        /// </summary>
        public string GetSqlCommand(Type modelType, string where, string orderBy, bool getValidValues = false)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.TableName = TableName;
                return DAO.GetSqlCommand(modelType, where, orderBy, getValidValues);
            }
            else
            {
                SqlDAO sqlDAO = new SqlDAO();
                return sqlDAO.GetSqlCommand(modelType, where, orderBy);
            }
        }
        #endregion

        #region GetNextCode
        /// <summary>
        /// Retorna o próximo código
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <returns>Código</returns>
        public string GetNextCode()
        {
            var t = GetNextCode(TableName, "Code", String.Empty);
            return t;
        }

        /// <summary>
        /// Retorna o próximo código
        /// </summary>
        /// <param name="tableName">Nome da tabela</param>
        /// <returns>Código</returns>
        public static string GetNextCode(string tableName)
        {
            var t = GetNextCode(tableName, "Code", String.Empty);
            return t;
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
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                return CrudB1DAO.GetNextCode(tableName, fieldName, where);
            }
            else
            {
                return SqlDAO.GetNextCode(tableName, fieldName, where);
            }
        }
        #endregion GetNextCode

        #region GetColumnValue
        /// <summary>
        /// Retorna valor da coluna de acordo com o select
        /// </summary>
        public static T GetColumnValue<T>(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                return CrudB1DAO.GetColumnValue<T>(sql);
            }
            else
            {
                throw new NotImplementedException();
            }
        }
        #endregion

        public List<string> FillStringList(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                return CrudB1DAO.FillStringList(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                return DAO.FillStringList(sql);
            }
        }

        #region FillModel
        /// <summary>
        /// Preenche model de acordo com Annotation HubModel
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
        /// Preenche a lista de model através da Annotation HubModel
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public List<T> FillModelList<T>(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.TableName = TableName;
                return DAO.FillModelList<T>(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                return DAO.FillModelList<T>(sql);
            }
        }
        #endregion FillModel

        #region FillModelAccordingTOSql
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
        /// Preenche a lista de model através do SQL
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="sql">Comando SQL</param>
        /// <returns>Lista de Model preenchido</returns>
        public List<T> FillModelListFromSql<T>(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.TableName = TableName;
                return DAO.FillModelListFromSql<T>(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                return DAO.FillListFromSql<T>(sql);
            }
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
        /// Retorna a quantidade de linhas de uma query
        /// </summary>
        /// <param name="sql">SELECT</param>
        /// <returns></returns>
        public int GetRowCount(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                return DAO.GetRowCount(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                return DAO.GetRowCount(sql);
            }
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
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.TableName = TableName;
                return DAO.Exists(returnColumn, where);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.TableName = TableName;
                throw new NotImplementedException();
            }
        }
        #endregion Exists

        public static void ExecuteNonQuery(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO.ExecuteNonQuery(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.ExecuteNonQuery(sql);
            }
        }

        public static object ExecuteScalar(string sql)
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                return CrudB1DAO.ExecuteScalar(sql);
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                return DAO.ExecuteScalar(sql);
            }
        }
        #endregion Util

        #region Transaction
        public void BeginTransaction()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.BeginTransaction();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.BeginTransaction();
            }
        }

        public void CommitTransaction()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.CommitTransaction();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.CommitTransaction();
            }
        }

        public void RollbackTransaction()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.RollbackTransaction();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.RollbackTransaction();
            }
        }

        public void InTransaction()
        {
            if (HubApp.AppType == AppTypeEnum.SBO)
            {
                CrudB1DAO DAO = new CrudB1DAO();
                DAO.RollbackTransaction();
            }
            else
            {
                SqlDAO DAO = new SqlDAO();
                DAO.RollbackTransaction();
            }
        }
        #endregion
    }
}
