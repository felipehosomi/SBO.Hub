using System;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Runtime.InteropServices;

namespace SBO.Hub.UI
{
    public static class ColumnExtensions
    {
        /// <summary>
        /// Preenche um Column com os dados de um Recordset.
        /// </summary>
        /// <param name="column">O Column a ser preenchido.</param>
        /// <param name="recordset">O Recordset com os dados.</param>
        /// <param name="fieldValue">O nome do Field que contém o valor para Column.</param>
        /// <param name="fieldDescription">O nome do Field que contém a descrição para o Column.</param>
        public static Recordset AddValuesFromRecordset(this Column column, Recordset recordset, string fieldValue = null,
                                                string fieldDescription = null)
        {
            if (column == null) throw new ArgumentNullException("column");
            if (recordset == null) throw new ArgumentNullException("recordset");

            recordset.MoveFirst();

            if (fieldDescription == null)
            {
                if (fieldValue == null)
                {
                    var fields = recordset.Fields;

                    if (fields.Count > 1)
                    {
                        fieldValue = fields.Item(0).Name;
                        fieldDescription = fields.Item(1).Name;
                    }
                    else
                    {
                        fieldValue = fieldDescription = fields.Item(0).Name;
                    }
                }
                else
                {
                    var fields = recordset.Fields;

                    if (fields.Count > 1)
                    {
                        fieldDescription = fields.Item(0).Name == fieldValue ? fields.Item(1).Name : fields.Item(0).Name;
                    }
                    else
                    {
                        fieldDescription = fieldValue;
                    }
                }
            }

            var validValues = column.ValidValues;

            while (!recordset.EoF)
            {
                var value = recordset.Fields.Item(fieldValue);
                var description = recordset.Fields.Item(fieldDescription);

                validValues.Add(value.Value.ToString(), description.Value.ToString());

                recordset.MoveNext();
            }

            return recordset;
        }

        public static void AddValidValues(this Column column, string tableName, string fieldName)
        {
            string sql = $@"SELECT 
	                            T1.FldValue,
	                            T1.Descr
                            FROM CUFD T0 
	                            INNER JOIN UFD1 T1 
		                            ON T0.TableID = T1.TableID AND T0.FieldID = T1.FieldID
                            WHERE T0.TableID = '{tableName}'
                            AND T0.AliasID = '{fieldName}'";
            AddValuesFromQuery(column, sql);
        }

        /// <summary>
        /// Preenche um Column com os dados de uma query.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="recordset">O Recordset a ser utilizado para obter os dados.</param>
        /// <param name="query">A query a ser utilizada para obter os dados.</param>
        /// <param name="fieldValue">O nome do Field que contém o valor para Column.</param>
        /// <param name="fieldDescription">O nome do Field que contém a descrição para o Column.</param>
        /// <param name="noLock">Indica se a query deve ser executada na forma de apenas leiura, isto é, se houver alguma escrita no momento não vai esperar a conclusão desta para continuar.</param>
        public static Recordset AddValuesFromQuery(this Column column, Recordset recordset, string query,
                                            string fieldValue = null, string fieldDescription = null, bool noLock = true)
        {
            if (recordset == null) throw new ArgumentNullException("recordset");
            if (query == null) throw new ArgumentNullException("query");

            recordset.DoQuery(noLock ? " SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED " + query : query);

            return AddValuesFromRecordset(column, recordset, fieldValue, fieldDescription);
        }

        /// <summary>
        /// Preenche um Column com os dados de uma query.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="company">O Company a ser utilizado para obter dos dados.</param>
        /// <param name="query">A query a ser utilizada para obter os dados.</param>
        /// <param name="fieldValue">O nome do Field que contém o valor para Column.</param>
        /// <param name="fieldDescription">O nome do Field que contém a descrição para o Column.</param>
        /// <param name="noLock">Indica se a query deve ser executada na forma de apenas leiura, isto é, se houver alguma escrita no momento não vai esperar a conclusão desta para continuar.</param>
        public static Recordset AddValuesFromQuery(this Column column, SAPbobsCOM.Company company, string query,
                                            string fieldValue = null, string fieldDescription = null, bool noLock = true)
        {
            if (company == null) throw new ArgumentNullException("company");

            var recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

            return AddValuesFromQuery(column, recordset, query, fieldValue, fieldDescription, noLock);
        }

        /// <summary>
        /// Preenche um Column com os dados de uma query.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="query">A query a ser utilizada para obter os dados.</param>
        /// <param name="fieldValue">O nome do Field que contém o valor para Column.</param>
        /// <param name="fieldDescription">O nome do Field que contém a descrição para o Column.</param>
        public static void AddValuesFromQuery(this Column column, string query, string fieldValue = null, string fieldDescription = null)
        {
            Recordset rst = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rst.DoQuery(query);

            if (String.IsNullOrEmpty(fieldValue))
            {
                fieldValue = rst.Fields.Item(0).Name;
            }

            if (String.IsNullOrEmpty(fieldDescription))
            {
                fieldDescription = rst.Fields.Count == 1 ? rst.Fields.Item(0).Name : rst.Fields.Item(1).Name;
            }

            while (!rst.EoF)
            {
                var value = rst.Fields.Item(fieldValue);
                var description = rst.Fields.Item(fieldDescription);

                column.ValidValues.Add(value.Value.ToString(), description.Value.ToString());

                rst.MoveNext();
            }

            Marshal.ReleaseComObject(rst);
            rst = null;
        }
    }
}
