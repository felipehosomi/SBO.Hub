using SBO.Hub.Attributes;

namespace SBO.Hub.Models
{
    public class DefaultMasterDataModel
    {
        [HubModel()]
        public string Code { get; set; }

        [HubModel()]
        public string Name { get; set; }
    }
}
