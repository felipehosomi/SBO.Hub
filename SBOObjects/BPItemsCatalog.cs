using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace SBO.Hub.SBOObjects
{
    public class BPItemsCatalog
    {
        public static string GetItemCode(string cardCode, string substitute)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = @"SELECT ItemCode FROM OSCN WHERE Substitute = '{0}' AND CardCode = '{1}'";
            sql = String.Format(sql, substitute, cardCode);
            rs.DoQuery(sql);

            string itemCode = String.Empty;
            if (rs.RecordCount > 0)
            {
                itemCode = rs.Fields.Item("ItemCode").Value.ToString();
            }

            Marshal.ReleaseComObject(rs);
            rs = null;

            return itemCode;
        }

        public static string GetSubstitute(string cardCode, string itemCode)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = @"SELECT Substitute FROM OSCN WHERE ItemCode = '{0}' AND CardCode = '{1}'";
            sql = String.Format(sql, itemCode, cardCode);
            rs.DoQuery(sql);

            string substitute = String.Empty;
            if (rs.RecordCount > 0)
            {
                substitute = rs.Fields.Item("Substitute").Value.ToString();
            }

            Marshal.ReleaseComObject(rs);
            rs = null;

            return substitute;
        }

        public static string GetItemFromBPCatalog(string cardCode, string itemCode)
        {
            Recordset rs = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = @"SELECT ItemName 
                            FROM OITM WITH(NOLOCK)
                            INNER JOIN OSCN WITH(NOLOCK)
                                ON OSCN.ItemCode = OITM.ItemCode
                            WHERE OSCN.CardCode = '{0}'
                            AND OSCN.ItemCode = '{1}'";

            rs.DoQuery(String.Format(sql, cardCode, itemCode));

            string itemName = String.Empty;

            if (rs.RecordCount > 0)
            {
                itemName = rs.Fields.Item(0).Value.ToString();
            }

            Marshal.ReleaseComObject(rs);
            rs = null;

            return itemName;
        }
    }
}
