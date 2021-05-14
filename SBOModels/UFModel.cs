using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class UFModel
    {
        [HubModel]
        public string Code { get; set; }

        [HubModel]
        public string Name { get; set; }
    }
}
