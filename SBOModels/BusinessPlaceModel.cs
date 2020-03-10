using SBO.Hub.Attributes;

namespace SBO.Hub.SBOModels
{
    public class BusinessPlaceModel
    {
        [HubModel]
        public int BPlId { get; set; }
        
        [HubModel]
        public string BPlName { get; set; }

        [HubModel(ColumnName = "TaxIdNum")]
        public string Cnpj { get; set; }

        [HubModel(ColumnName = "AddtnlId")]
        public string InscMunicipal { get; set; }
    }
}
