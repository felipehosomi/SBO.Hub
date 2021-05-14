using SBO.Hub.Enums;
using System;

namespace SBO.Hub.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public class HubModelAttribute : Attribute
    {
        public int Index { get; set; }

        /// <summary>
        /// Nome da Coluna na tabela
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// É chave primária
        /// </summary>
        public bool IsPK { get; set; }

        /// <summary>
        /// Name da Tabela
        /// </summary>
        public string TableName { get; set; }

        /// <summary>
        /// Nome do campo na tela - Utilizado em Grid/Matrix/DataTable
        /// </summary>
        public string UIFieldName { get; set; }

        /// <summary>
        /// Ignorar campo na tela - Utilizado em Grid/Matrix/DataTable
        /// </summary>
        public bool UIIgnore { get; set; }

        /// <summary>
        /// Descricao do campo -- Utilizado caso queira alterar o nome da coluna no comando SQL
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Nome do UserDataSource do campo
        /// </summary>
        public string UserDataSource { get; set; }

        /// <summary>
        /// Campo obrigatório
        /// </summary>
        public bool MandatoryYN { get; set; }
        
        /// <summary>
        /// Preenche campo (Banco)?
        /// </summary>
        public bool FillOnSelect { get; set; }

        /// <summary>
        /// Preenche campo (Tela)?
        /// </summary>
        public bool AutoFill { get; set; }

        /// <summary>
        /// Campo para ser tratado no banco de dados?
        /// </summary>
        public bool DataBaseFieldYN { get; set; }

        public UIFieldTypeEnum UIFieldType { get; set; }

        public HubModelAttribute()
        {
            DataBaseFieldYN = true;
            AutoFill = true;
            FillOnSelect = true;
            UIIgnore = false;
        }
    }
}
