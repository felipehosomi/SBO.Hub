using System;

namespace SBO.Hub.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class FormAttribute : Attribute
    {
        public string FormId { get; set; }

        public FormAttribute(string formId)
        {
            this.FormId = formId;
        }
    }
}
