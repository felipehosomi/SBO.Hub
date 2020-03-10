using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class CityModel
    {
        [HubModel()]
        public int AbsId { get; set; }

        [HubModel()]
        public string Code { get; set; }

        [HubModel()]
        public string State { get; set; }

        [HubModel()]
        public string Name { get; set; }

        [HubModel()]
        public string TaxZone { get; set; }

        [HubModel()]
        public string IbgeCode { get; set; }

        [HubModel()]
        public string GiaCode { get; set; }
    }
}
