using SBO.Hub.Enums;

namespace SBO.Hub
{
    public class HubApp
    {
        public static string ServerName { get; set; }

        public static string DatabaseName { get; set; }

        public static string DBUserName { get; set; }

        public static string DBPassword { get; set; }

        public static AppTypeEnum AppType { get; set; }

        public static DatabaseTypeEnum DatabaseType { get; set; }

        public static void FillConnectionParameters()
        {
            DatabaseName = System.Configuration.ConfigurationManager.AppSettings["Database"];
            ServerName = System.Configuration.ConfigurationManager.AppSettings["Server"];
            DBUserName = System.Configuration.ConfigurationManager.AppSettings["DBUserName"];
            ServerName = System.Configuration.ConfigurationManager.AppSettings["DBPassword"];
        }

        public static void FillConnectionParameters(string database, string server, string dbUser, string dbPassword)
        {
            DatabaseName = database;
            ServerName = server;
            DBUserName = dbUser;
            ServerName = dbPassword;
        }
    }
}
