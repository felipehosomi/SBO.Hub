using SAPbouiCOM;
using System.Xml;

namespace SBO.Hub.Util
{
    public class SBOXmlDocument
    {
        public static string GetXmlField(BusinessObjectInfo businessObjectInfo, string field = "DocEntry")
        {

            XmlDocument oXmlDoc = new XmlDocument();
            oXmlDoc.LoadXml(businessObjectInfo.ObjectKey);

            XmlNodeList oXmlNodeList = oXmlDoc.GetElementsByTagName(field);
            string ValorChave = oXmlNodeList[0].InnerText;

            return ValorChave;
        }

    }
}
