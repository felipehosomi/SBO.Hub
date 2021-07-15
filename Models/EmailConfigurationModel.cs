using System.Collections.Generic;

namespace SBO.Hub.Models
{
    public class EmailConfigurationModel
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Server { get; set; }
        public int Port { get; set; }
        public string SSL { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public List<string> MailTo { get; set; }
    }
}
