using System;

namespace SBO.Hub.Models
{
    public class CertificateModel
    {
        public string Name { get; set; }
        public string SerialNumber { get; set; }
        public DateTime ExpirationDate { get; set; }
    }
}
