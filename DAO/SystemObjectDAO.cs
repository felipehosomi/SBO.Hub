using System;
using System.Collections.Generic;

namespace SBO.Hub.DAO
{
    public class SystemObjectDAO
    {
        #region Properties
        private CrudDAO CrudDAO;
        private string tableName;
        #endregion Properties

        #region Constructor
        public SystemObjectDAO(string tableName)
        {
            CrudDAO = new CrudDAO(tableName);
            this.tableName = tableName;
        }
        #endregion Constructor

        #region Retrieve
        public virtual string GetSqlCommand(Type modelType, string where, string orderBy, bool getValidValues)
        {
            return CrudDAO.GetSqlCommand(modelType, where, orderBy, getValidValues);
        }

        public virtual string RetrieveSqlModel(Type modelType, string where, bool getValidValues)
        {
            return this.GetSqlCommand(modelType, where, String.Empty, getValidValues);
        }

        public virtual string RetrieveSqlModel(Type modelType, bool getValidValues)
        {
            return this.GetSqlCommand(modelType, String.Empty, String.Empty, getValidValues);
        }

        public virtual T RetrieveModel<T>(string where)
        {
            return CrudDAO.RetrieveModel<T>(where);
        }

        public virtual List<T> RetrieveModelList<T>(string where)
        {
            return CrudDAO.RetrieveModelList<T>(where);
        }

        public virtual List<T> RetrieveModelList<T>(string where, string orderBy)
        {
            return CrudDAO.RetrieveModelList<T>(where, orderBy);
        }
        #endregion

        #region Util
        public virtual string Exists(string where)
        {
            return CrudDAO.Exists(where);
        }

        public virtual string Exists(string returnColumn, string where)
        {
            return CrudDAO.Exists(returnColumn, where);
        }

        public virtual List<T> FillModelList<T>(string sql)
        {
            return CrudDAO.FillModelList<T>(sql);
        }

        public string GetNextCode()
        {
            return CrudDAO.GetNextCode(tableName);
        }

        public string GetNextCode(string fieldName)
        {
            return CrudDAO.GetNextCode(fieldName, tableName);
        }
        #endregion
    }
}
