using SBO.Hub.Enums;
using System;

namespace SBO.Hub.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class TextFileAttribute : Attribute
    {
        private int position;
        public int Position
        {
            get
            {
                return position - 1;
            }
            set
            {
                position = value;
            }
        }

        public int Size { get; set; }

        public int DecimalPlaces { get; set; } = 2;

        public string DecimalSeparator { get; set; }

        public string Format { get; set; }

        public PaddingTypeEnum PaddingType { get; set; } = PaddingTypeEnum.NotSet;

        public string PaddingChar { get; set; }

        public bool OnylNumeric { get; set; }
    }
}
