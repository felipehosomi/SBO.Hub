using SAPbobsCOM;
using System;
using System.Collections.Generic;

namespace SBO.Hub.DAO
{
    public class UserDefinedObjectDAO
    {
        #region Properties
        private CrudDAO crudDAO;
        private string tableName;
        #endregion Properties

        #region Constructor
        public UserDefinedObjectDAO(string tableName)
        {
            crudDAO = new CrudDAO(tableName);
            this.tableName = tableName;
        }

        public UserDefinedObjectDAO(string tableName, BoUTBTableType userTableType)
        {
            crudDAO = new CrudDAO(tableName);
            crudDAO.UserTableType = userTableType;
            this.tableName = tableName;
        }
        #endregion Constructor

        #region CRUD
        public virtual string GetSqlCommand(Type modelType, string where, string orderBy, bool getValidValues)
        {
            return crudDAO.GetSqlCommand(modelType, where, orderBy, getValidValues);
        }

        public virtual string RetrieveSqlModel(Type modelType, string where, bool getValidValues)
        {
            return this.GetSqlCommand(modelType, where, String.Empty, getValidValues);
        }

        public virtual string GetSqlCommand(Type modelType, bool getValidValues)
        {
            return this.GetSqlCommand(modelType, String.Empty, String.Empty, getValidValues);
        }

        public virtual void CreateModel(object model)
        {
            crudDAO.Model = model;
            crudDAO.CreateModel();
        }

        public virtual T RetrieveModel<T>(string where)
        {
            return crudDAO.RetrieveModel<T>(where);
        }

        public virtual List<T> RetrieveModelList<T>(string where)
        {
            return crudDAO.RetrieveModelList<T>(where);
        }

        public virtual List<T> RetrieveModelList<T>(string where, string orderBy)
        {
            return crudDAO.RetrieveModelList<T>(where, orderBy);
        }

        public virtual void UpdateModel(object model)
        {
            crudDAO.Model = model;
            crudDAO.UpdateModel();
        }

        public virtual void UpdateModel(object model, string where)
        {
            crudDAO.Model = model;
            crudDAO.UpdateModel(where);
        }

        public virtual void DeleteModel(string where)
        {
            crudDAO.DeleteModel(tableName, where);
        }

        public virtual void DeleteModelByCode(string code)
        {
            crudDAO.DeleteModelByCode(tableName, code);
        }

        #endregion

        #region Util
        public virtual string Exists(string where)
        {
            return crudDAO.Exists(where);
        }

        public virtual string Exists(string returnColumn, string where)
        {
            return crudDAO.Exists(returnColumn, where);
        }

        public virtual T FillModel<T>(string sql)
        {
            return crudDAO.FillModel<T>(sql);
        }

        public virtual List<T> FillModelList<T>(string sql)
        {
            return crudDAO.FillModelList<T>(sql);
        }

        public string GetNextCode()
        {
            return CrudDAO.GetNextCode(tableName);
        }

        public string GetNextCode(string fieldName)
        {
            return CrudDAO.GetNextCode(tableName, fieldName);
        }

        public string GetNextCode(string fieldName, string where)
        {
            return CrudDAO.GetNextCode(tableName, fieldName, where);
        }
        #endregion
    }
}
