using SAPbobsCOM;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace SBO.Hub.Helpers
{
    public class CrystalReports
    {
        public static string GetReportLayout(string reportName)
        {
            string format = "SELECT \"TypeCode\" FROM RDOC WHERE \"DocName\" = '{0}'";
            format = string.Format(format, reportName);
            Recordset businessObject = (Recordset)SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            businessObject.DoQuery(format);
            string str2 = businessObject.Fields.Item("TypeCode").Value.ToString();
            Marshal.ReleaseComObject(businessObject);
            businessObject = null;
            return str2;
        }

        public static bool GetReportMenu(string docName)
        {
            string ssql = @"select * from rtyp where ""ADD_NAME"" = '{0}'";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SBOApp.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(string.Format(ssql, docName));
            int rss = rs.RecordCount;

            Marshal.ReleaseComObject(rs);
            GC.Collect();

            if (rss != 0)
            {
                return false;
            }
            return true;

        }

        #region Report
        /// <summary>
        /// Adiciona report no form
        /// </summary>
        /// <param name="reportFile">Nome do arquivo - Ex: report.rpt</param>
        /// <param name="reportName">Ex: Relatório</param>
        /// <param name="typeName">SAP AddOn Aços-Continente</param>
        /// <param name="addOnName">Relatório</param>
        /// <param name="addOnForm">410000100</param>
        /// <param name="menuId">10002</param>
        /// <returns></returns>
        public static string AddReportToForm(string reportFile, string reportName, string typeName, string addOnName, string addOnForm, string menuId)
        {
            CompanyService service = SBOApp.Company.GetCompanyService();
            ReportTypesService rptTypeService = (ReportTypesService)service.GetBusinessService(ServiceTypes.ReportTypesService);
            ReportType newType = (ReportType)rptTypeService.GetDataInterface(ReportTypesServiceDataInterfaces.rtsReportType);

            newType.TypeName = typeName;
            newType.AddonName = addOnName;
            newType.AddonFormType = addOnForm;
            newType.MenuID = menuId;

            bool createReport = CrystalReports.GetReportMenu(newType.AddonName);
            if (createReport)
            {
                ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);

                ReportLayoutsService rptService = (ReportLayoutsService)
                service.GetBusinessService(ServiceTypes.ReportLayoutsService);
                ReportLayout newReport = (ReportLayout)rptService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
                newReport.Author = SBOApp.Company.UserName;
                newReport.Category = ReportLayoutCategoryEnum.rlcCrystal;
                newReport.Name = reportName;
                newReport.TypeCode = newTypeParam.TypeCode;
                ReportLayoutParams newReportParam = rptService.AddReportLayout(newReport);

                newType = rptTypeService.GetReportType(newTypeParam);
                newType.DefaultReportLayout = newReportParam.LayoutCode;
                rptTypeService.UpdateReportType(newType);

                BlobParams oBlobParams = (BlobParams)
                service.GetDataInterface(CompanyServiceDataInterfaces.csdiBlobParams);
                oBlobParams.Table = "RDOC";
                oBlobParams.Field = "Template";
                BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
                oKeySegment.Name = "DocCode";
                oKeySegment.Value = newReportParam.LayoutCode;

                string appPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                if (!appPath.EndsWith(@"\")) appPath += @"\";

                FileStream oFile = new FileStream(String.Format(@"{0}Reports\{1}", appPath, reportFile), System.IO.FileMode.Open);
                int fileSize = (int)oFile.Length;
                byte[] buf = new byte[fileSize];
                oFile.Read(buf, 0, fileSize);
                //oFile.Dispose();
                oFile.Close();
                Blob oBlob = (Blob)service.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob);
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
                service.SetBlob(oBlobParams, oBlob);
                return newType.TypeCode;
            }
            else
            {
                return CrystalReports.GetReportLayout(reportName);
            }
        }
        #endregion
    }
}
