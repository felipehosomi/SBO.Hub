using SAPbobsCOM;
using SBO.Hub.DAO;
using System;
using System.Runtime.InteropServices;

namespace SBO.Hub.Controllers
{
    public class UserController : SystemObjectDAO
    {
        public UserController()
            : base("OUSR")
        { }

        public static int GetUserId(string userCode)
        {
            UserController userController = new UserController();
            string userId = userController.Exists("USERID", String.Format("USER_CODE = '{0}'", userCode));
            
            return Convert.ToInt32(userId);            
        }

        public static object GetField(string fieldName)
        {
            Recordset rst = (Recordset)SBOApp.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                rst.DoQuery(String.Format("SELECT OUSR.{0} FROM OUSR WHERE OUSR.USER_CODE = '{1}'", fieldName, SBOApp.Company.UserName));

                if (rst.RecordCount > 0)
                {
                    return rst.Fields.Item(0).Value.ToString();
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(rst);
                rst = null;
            }
        }
    }
}
