using SAPbobsCOM;
using System;
using System.Data;

namespace SBO.Hub.SBOHelpers
{
    public static class RecordsetExtensions
    {
        public static DataTable ToDataTable(this Recordset rst, string tableName = "")
        {
            DataTable dataTable = new DataTable(tableName);
            for (int i = 0; i < rst.Fields.Count; i++)
            {
                Type type = typeof(string);
                switch (rst.Fields.Item(i).Type)
                {
                    case BoFieldTypes.db_Alpha:
                    case BoFieldTypes.db_Memo:
                        type = typeof(string);
                        break;
                    case BoFieldTypes.db_Numeric:
                        type = typeof(int);
                        break;
                    case BoFieldTypes.db_Date:
                        type = typeof(DateTime);
                        break;
                    case BoFieldTypes.db_Float:
                        type = typeof(double);
                        break;
                    default:
                        break;
                }

                dataTable.Columns.Add(rst.Fields.Item(i).Name, type);
            }

            rst.MoveFirst();
            while (!rst.EoF)
            {
                DataRow row = dataTable.NewRow();
                for (int i = 0; i < rst.Fields.Count; i++)
                {
                    row[i] = rst.Fields.Item(i).Value;
                }
                dataTable.Rows.Add(row);
                rst.MoveNext();
            }
            
            return dataTable;
        }
    }
}
