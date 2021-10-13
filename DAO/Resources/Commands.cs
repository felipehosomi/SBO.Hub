using System.Resources;

namespace SBO.Hub.DAO.Resources
{
    public class Commands
    {
        public static ResourceManager Resource;

        public static void SetResourceManager()
        {
            if (System.Configuration.ConfigurationManager.AppSettings["ServerType"] == "9")
            {
                Resource = new ResourceManager("SBO.Hub.DAO.Resources.Hana", typeof(Hana).Assembly);
            }
            else
            {
                Resource = new ResourceManager("SBO.Hub.DAO.Resources.SQL", typeof(SQL).Assembly);
            }
        }
    }
}
