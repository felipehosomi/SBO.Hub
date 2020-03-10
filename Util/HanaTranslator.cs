using System;
using Translator;

namespace SBO.Hub.Util
{
    public class HanaTranslator
    {
        public static TranslatorTool Translator { get; private set; }

        public static string Translate(string sql)
        {
            int count;
            int errCount;
            try
            {
                if (Translator == null)
                {
                    Translator = new TranslatorTool();
                }
                string hana = Translator.TranslateQuery(sql, out count, out errCount);
                if (errCount == 0)
                {
                    sql = hana.Substring(0, hana.Length - 3);
                }
                return sql;
            }
            catch (Exception ex)
            {

            }
            return sql;

        }
    }
}
