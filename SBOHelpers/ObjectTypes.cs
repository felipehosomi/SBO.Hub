using SBO.Hub.Enums;
using System;

namespace SBO.Hub.SBOHelpers
{
    public class ObjectTypes
    {
        public static string GetTable(ObjectTypeEnum objectTypeEnum)
        {
            switch (objectTypeEnum)
            {
                case ObjectTypeEnum.NotaFiscalSaida:
                    return "INV";
                case ObjectTypeEnum.DevNotaFiscalSaida:
                    return "RIN";
                case ObjectTypeEnum.Entrega:
                    return "DLN";
                case ObjectTypeEnum.Devolucao:
                    return "RDN";
                case ObjectTypeEnum.PedidoVenda:
                    return "RDR";
                case ObjectTypeEnum.NotaFiscalEntrada:
                    return "PCH";
                case ObjectTypeEnum.DevNotaFiscalEntrada:
                    return "RPC";
                case ObjectTypeEnum.RecebimentoMercadorias:
                    return "PDN";
                case ObjectTypeEnum.DevolucaoMercadorias:
                    return "RPD";
                case ObjectTypeEnum.PedidoCompra:
                    return "POR";
                case ObjectTypeEnum.Cotacao:
                    return "QUT";
                case ObjectTypeEnum.SolicitacaoCompra:
                    return "PRQ";
            }
            throw new Exception("Objeto não encontrado - Adicione no framework se estiver disponível!");
        }

        public static string GetDocumentTable(ObjectTypeEnum objectTypeEnum)
        {
            return "O" + GetTable(objectTypeEnum);
        }
    }
}
