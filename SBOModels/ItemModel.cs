using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class ItemModel
    {
        [HubModel(ColumnName = "ItemCode")]
        public string ItemCode { get; set; }

        [HubModel(ColumnName = "ItemName")]
        public string ItemName { get; set; }
    }
}
