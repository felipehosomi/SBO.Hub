using SBO.Hub.Attributes;
using SBO.Hub.Controllers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Text;

namespace SBO.Hub.DAO
{
    public class SqlDAO : IDAO
    {
        private static SqlConnection Connection = new SqlConnection();
        private SqlDataAdapter DataAdapter = new SqlDataAdapter();
        private SqlDataReader DataReader;
        private SqlCommand Command;
        private static SqlTransaction Transaction;

        public string TableName { get; set; }
        public object Model { get; set; }

        private static string ConnectionString;
        public static string Database { get; set; }

        public SqlDAO()
        {
            if (String.IsNullOrEmpty(ConnectionString))
            {
                try
                {
                    Database = HubApp.DatabaseName;
                    string server = HubApp.ServerName;
                    string dbUser = HubApp.DBUserName;
                    string dbPassword = HubApp.DBPassword;

                    SetConnectionString(server, Database, dbUser, dbPassword);
                }
                catch { }
            }
        }

        public SqlDAO(string connectionString)
        {
            ConnectionString = connectionString;
        }

        public SqlDAO(string database, string server, string dbUser, string dbPassword)
        {
            SetConnectionString(server, Database, dbUser, dbPassword);
        }

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
            string sql = String.Format(" SELECT ISNULL(MAX(CAST({0} AS BIGINT)), 0) + 1 FROM [{1}] ", fieldName, tableName);

            if (!String.IsNullOrEmpty(where))
            {
                sql += String.Format(" WHERE {0} ", where);
            }

            SqlDAO dao = new SqlDAO();
            object code = dao.ExecuteScalar(sql);

            if (code != null)
            {
                return Convert.ToInt32(code).ToString();
            }
            else
            {
                return String.Empty;
            }
        }
        #endregion GetNextCode

        #region CreateModel
        /// <summary>
        /// Salva o model no banco de dados
        /// </summary>
        public void CreateModel()
        {
            try
            {
                HubModelAttribute hubModel;

                string sqlColumns = String.Empty;
                string sqlValues = String.Empty;

                List<Type> typesList = new List<Type>();
                List<object> valuesList = new List<object>();

                PropertyInfo propCode = Model.GetType().GetProperty("Code");
                if (propCode != null)
                {
                    if (propCode.GetValue(Model, null) == null)
                    {
                        propCode.SetValue(Model, GetNextCode(TableName).PadLeft(8, '0'));
                    }
                }

                //Dictionary<Type, object> values = new Dictionary<Type, object>();
                //List<object> values = new List<object>();
                // Percorre as propriedades do Model
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    try
                    {
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

                                typesList.Add(property.PropertyType);
                                valuesList.Add(property.GetValue(Model, null));

                                if (String.IsNullOrEmpty(hubModel.ColumnName))
                                {
                                    hubModel.ColumnName = property.Name;
                                }

                                sqlColumns += String.Format(", {0}", hubModel.ColumnName);
                                break;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception(String.Format("Erro ao setar propriedade {0}: {1}", property.Name, e));
                    }
                }
                if (String.IsNullOrEmpty(sqlColumns))
                {
                    throw new Exception("Nenhuma coluna informada. Informe a propriedade ColumnName no Model");
                }

                for (int i = 0; i < typesList.Count; i++)
                {
                    if (valuesList[i] == null || valuesList[i] == DBNull.Value)
                    {
                        sqlValues += ", NULL";
                    }
                    else if (typesList[i] == typeof(string) || typesList[i] == typeof(String) || typesList[i] == typeof(char))
                    {
                        sqlValues += String.Format(", '{0}'", valuesList[i].ToString().Replace("'", "''"));
                    }
                    else if (typesList[i] == typeof(DateTime) || typesList[i] == typeof(Nullable<DateTime>))
                    {
                        sqlValues += String.Format(", CONVERT(DATETIME, '{0}') ", ((DateTime)valuesList[i]).ToString("yyyy-MM-ddTHH:mm:ss"));
                    }
                    else
                    {
                        sqlValues += String.Format(", '{0}' ", valuesList[i].ToString().Replace(",", "."));
                    }
                }

                string sql = String.Format(" INSERT INTO [{0}] ({1}) VALUES ({2}) ", TableName, sqlColumns.Substring(2), sqlValues.Substring(2));

                this.ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region UpdateModel
        public void UpdateModel()
        {
            try
            {
                HubModelAttribute hubModel;

                string sqlWhere = String.Empty;
                string sqlValues = String.Empty;
                object value;

                Dictionary<Type, object> values = new Dictionary<Type, object>();
                //List<object> values = new List<object>();
                // Percorre as propriedades do Model
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    try
                    {
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
                                {
                                    hubModel.ColumnName = property.Name;
                                }

                                value = property.GetValue(Model, null);
                                if (hubModel.IsPK || property.Name == "Code")
                                {
                                    if (value == null)
                                    {
                                        sqlWhere += String.Format(" AND {0} = NULL", hubModel.ColumnName);
                                    }
                                    else if (property.PropertyType == typeof(string))
                                    {
                                        sqlWhere += String.Format(" AND {0} = '{1}'", hubModel.ColumnName, value);
                                    }
                                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                                    {
                                        sqlWhere += String.Format(" AND {0} = CONVERT(DATETIME, '{1}')", hubModel.ColumnName, Convert.ToDateTime(value).ToString("yyyyMMdd"));
                                    }
                                    else
                                    {
                                        sqlWhere += String.Format(" AND {0} = {1}", hubModel.ColumnName, value.ToString().Replace(",", "."));
                                    }
                                }
                                else
                                {
                                    if (value == null)
                                    {
                                        sqlValues += String.Format(", {0} = NULL", hubModel.ColumnName);
                                    }
                                    else if (property.PropertyType == typeof(string))
                                    {
                                        sqlValues += String.Format(", {0} = '{1}'", hubModel.ColumnName, value.ToString().Replace("'", "''"));
                                    }
                                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                                    {
                                        sqlValues += String.Format(", {0} = CONVERT(DATETIME, '{1}')", hubModel.ColumnName, Convert.ToDateTime(value).ToString("yyyyMMdd"));
                                    }
                                    else
                                    {
                                        sqlValues += String.Format(", {0} = {1}", hubModel.ColumnName, value.ToString().Replace(",", "."));
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

                if (String.IsNullOrEmpty(sqlWhere))
                {
                    throw new Exception("Nenhuma coluna PK informada. Informe a propriedade IsPK no Model ou crie um campo chamado 'Code' (será utilizado como PK por default)");
                }
                if (String.IsNullOrEmpty(sqlValues))
                {
                    throw new Exception("Nenhuma coluna informada para atualizar. Informe a propriedade ColumnName no Model");
                }

                string sql = String.Format(" UPDATE [{0}] SET {1} WHERE {2} ", TableName, sqlValues.Substring(2), sqlWhere.Substring(4));
                this.ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        public string GetConnectedServer()
        {
            return Connection.DataSource;
        }

        public static void Connect()
        {
            if (Connection == null || Connection.ConnectionString != ConnectionString)
            {
                Connection = new SqlConnection();
            }

            if (Connection.State == ConnectionState.Broken || Connection.State == ConnectionState.Closed)
            {
                try
                {
                    Connection.ConnectionString = ConnectionString;
                    Connection.Open();
                }
                catch (Exception ex)
                {
                    throw new Exception("Erro ao conectar SQL: " + ex.Message);
                }
            }
        }

        public void Close()
        {
            if (Connection.State == ConnectionState.Open || Connection.State == ConnectionState.Executing || Connection.State == ConnectionState.Fetching)
            {
                Connection.Close();
                Connection.Dispose();
                Connection = null;
            }
        }

        public SqlDataReader ExecuteReader(string sql)
        {
            try
            {
                Connect();
                this.Command = new SqlCommand(sql, Connection);
                this.Command.CommandTimeout = 120;
                this.Command.Transaction = Transaction;
                this.DataReader = Command.ExecuteReader();
                return this.DataReader;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao executar SqlDataReader: " + ex.Message);
            }
        }

        public object ExecuteScalar(string sql)
        {
            try
            {
                Connect();
                this.Command = new SqlCommand(sql, Connection);
                this.Command.CommandTimeout = 120;
                this.Command.Transaction = Transaction;
                return this.Command.ExecuteScalar();
            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao executar ExecuteScalar: " + ex.Message);

            }
        }

        public void ExecuteNonQuery(string sql)
        {
            try
            {
                Connect();
                this.Command = new SqlCommand(sql, Connection);
                this.Command.CommandTimeout = 120;
                this.Command.Transaction = Transaction;
                this.Command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao executar ExecuteNonQuery: " + ex.Message);
            }
        }

        public DataTable FillDataTable(string sql)
        {
            try
            {
                Connect();
                this.Command = new SqlCommand(sql, Connection);
                this.Command.CommandTimeout = 120;
                this.Command.CommandType = CommandType.Text;
                this.Command.CommandText = sql;
                DataAdapter.SelectCommand = this.Command;

                DataTable dtb = new DataTable();

                DataAdapter.Fill(dtb);
                DataAdapter.Dispose();
                return dtb;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao executar FillDataTable: " + ex.Message);
            }
        }

        public void BeginTransaction()
        {
            Connect();
            Transaction = Connection.BeginTransaction();
        }

        public void RollbackTransaction()
        {
            if (Transaction.Connection != null)
            {
                Transaction.Rollback();
            }
        }

        public void CommitTransaction()
        {
            if (Transaction.Connection != null)
            {
                Transaction.Commit();
            }
        }

        public static void SetConnectionString(string serverName, string dataBaseName, string userName, string userPassword)
        {
            ConnectionString = String.Format(@" data source={0};initial catalog={1};persist security info=True;user id={2};password={3};",
                                                serverName,
                                                dataBaseName,
                                                userName,
                                                userPassword);
        }

        public T FillModel<T>(string sql)
        {
            List<T> modelList = this.FillModelList<T>(sql);
            if (modelList.Count > 0)
            {
                return modelList[0];
            }
            else
            {
                return Activator.CreateInstance<T>();
            }
        }

        public int GetRowCount(string sql)
        {
            int recordCount = 0;
            using (SqlDataReader dr = this.ExecuteReader(sql))
            {
                DataTable dt = new DataTable();
                dt.Load(dr);
                recordCount = dt.Rows.Count;
            }

            return recordCount;
        }

        public List<string> FillStringList(string sql)
        {

            List<string> list = new List<string>();
            using (SqlDataReader dr = this.ExecuteReader(sql))
            {
                while (dr.Read())
                {
                    if (!dr.IsDBNull(0))
                    {
                        list.Add(dr.GetValue(0).ToString());
                    }
                    else
                    {
                        list.Add(String.Empty);
                    }
                }
            }
            return list;
        }

        public List<T> FillModelList<T>(string sql)
        {
            sql = sql.Replace("\"", "");

            List<T> modelList = new List<T>();
            T model;
            HubModelAttribute hubModel;
            try
            {
                using (SqlDataReader dr = this.ExecuteReader(sql))
                {
                    while (dr.Read())
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
                                        {
                                            hubModel.ColumnName = property.Name;
                                        }
                                        if (!hubModel.FillOnSelect)
                                        {
                                            break;
                                        }

                                        int index = dr.GetOrdinal(hubModel.ColumnName);
                                        if (!dr.IsDBNull(index))
                                        {
                                            Type dbType = dr.GetFieldType(index);

                                            if (dbType == typeof(decimal) && property.PropertyType == typeof(double))
                                            {
                                                property.SetValue(model, Convert.ToDouble(dr.GetValue(index)), null);
                                            }
                                            else
                                            {
                                                property.SetValue(model, dr.GetValue(index), null);
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                throw new Exception(String.Format("Erro ao setar propriedade {0}: {1}", property.Name, e.Message));
                            }
                        }
                        modelList.Add(model);
                    }
                }
            }
            catch (Exception e)
            {

            }
            return modelList;
        }

        public T FillModelFromSql<T>(string sql)
        {
            List<T> modelList = this.FillListFromSql<T>(sql);
            if (modelList.Count > 0)
            {
                return modelList[0];
            }
            else
            {
                return Activator.CreateInstance<T>();
            }
        }

        public List<T> FillListFromSql<T>(string sql)
        {
            List<T> modelList = new List<T>();
            T model;
            using (SqlDataReader dr = this.ExecuteReader(sql))
            {
                while (dr.Read())
                {
                    // Cria nova instância do model
                    model = Activator.CreateInstance<T>();

                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        PropertyInfo property = model.GetType().GetProperty(dr.GetName(i));
                        if (property != null && !dr.IsDBNull(i))
                        {
                            try
                            {
                                property.SetValue(model, dr.GetValue(i), null);
                            }
                            catch
                            {
                                if (property.PropertyType == typeof(double))
                                {
                                    property.SetValue(model, Convert.ToDouble(dr.GetValue(i)), null);
                                }
                            }
                        }
                    }
                    modelList.Add(model);
                }
            }
            return modelList;
        }

        public void DeleteModel()
        {
            try
            {
                HubModelAttribute hubModel;

                string where = String.Empty;
                object value;

                Dictionary<Type, object> values = new Dictionary<Type, object>();
                //List<object> values = new List<object>();
                // Percorre as propriedades do Model
                foreach (PropertyInfo property in Model.GetType().GetProperties())
                {
                    try
                    {
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
                                {
                                    hubModel.ColumnName = property.Name;
                                }

                                value = property.GetValue(Model, null);
                                if (hubModel.IsPK || property.Name == "Code")
                                {
                                    if (value == null)
                                    {
                                        where += String.Format(" AND {0} = NULL", hubModel.ColumnName);
                                    }
                                    else if (property.PropertyType == typeof(string))
                                    {
                                        where += String.Format(" AND {0} = '{1}'", hubModel.ColumnName, value);
                                    }
                                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                                    {
                                        where += String.Format(" AND {0} = CONVERT(DATETIME, '{1}')", hubModel.ColumnName, Convert.ToDateTime(value).ToString("yyyyMMdd"));
                                    }
                                    else
                                    {
                                        where += String.Format(" AND {0} = {1}", hubModel.ColumnName, value.ToString().Replace(",", "."));
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

                if (String.IsNullOrEmpty(where))
                {
                    throw new Exception("Nenhuma coluna PK informada. Informe a propriedade IsPK no Model ou crie um campo chamado 'Code' (será utilizado como PK por default)");
                }

                string sql = String.Format(" DELETE FROM [{0}] WHERE {1} ", TableName, where.Substring(4));
                this.ExecuteNonQuery(sql);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool Exists(string where)
        {
            where = where.Replace("\"", "");
            string command = $"SELECT 1 FROM [{Database}].[dbo].[{TableName}] WHERE {where} ";
            return this.HasRows(command);
        }

        public bool HasRows(string command)
        {
            bool hasRows;
            using (SqlDataReader dr = this.ExecuteReader(command))
            {
                hasRows = dr.HasRows;
            }

            return hasRows;
        }

        public T RetrieveModel<T>(string where = "", string orderBy = "")
        {
            List<T> modelList = this.RetrieveModelList<T>(where, orderBy);
            if (modelList.Count > 0)
                return modelList[0];
            else
                return Activator.CreateInstance<T>();
        }

        public List<T> RetrieveModelList<T>(string where = "", string orderBy = "")
        {
            string sql = this.GetSqlCommand(typeof(T), where, orderBy);
            return FillModelList<T>(sql);
        }
        public string GetSqlCommand(Type modelType, string where, string orderBy)
        {
            StringBuilder command = new StringBuilder();
            // Inicia o SELECT
            command.Append(" SELECT ");
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
                        fields += String.Format(", [{0}]", hubModel.ColumnName);
                    }
                    break;
                }
            }

            if (String.IsNullOrEmpty(fields))
            {
                throw new Exception("Nenhuma propriedade do tipo hubModel encontrada no Model");
            }

            // Campos a serem retornados
            command.Append(fields.Substring(1));

            command.AppendFormat(" FROM [{0}] ", TableName);


            // Condição WHERE
            if (!String.IsNullOrEmpty(where))
            {
                command.AppendFormat(" WHERE {0} ", where);
            }

            // Condição ORDER BY
            if (!String.IsNullOrEmpty(orderBy))
            {
                command.AppendFormat(" ORDER BY {0} ", orderBy);
            }

            command = command.Replace("\"", String.Empty);
            return command.ToString();
        }

    }
}
