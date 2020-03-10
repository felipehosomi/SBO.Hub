using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class WarehouseModel
    {
        [HubModel]
        public string WhsCode { get; set; }

        [HubModel]
        public string WhsName { get; set; }
    }
}
