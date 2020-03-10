using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class CountryModel
    {
        [HubModel]
        public string Code { get; set; }

        [HubModel]
        public string Name { get; set; }
    }
}
