using SAPbouiCOM;
using SBO.Hub.Attributes;
using SBO.Hub.Controllers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace SBO.Hub.Forms
{
    public class BaseForm : SBO.Hub.Forms.IForm
    {
        #region Properties

        private static readonly Dictionary<Type, string> _formId;
        private static readonly Dictionary<Type, int> _formCount;
        private StringReader _srdSrfFile;
        private string _viewsFolder;
        private XmlDocument _formXml;

        public ItemEvent ItemEventInfo { get; set; }
        public BusinessObjectInfo BusinessObjectInfo { get; set; }
        public ContextMenuInfo ContextMenuInfo { get; set; }
        public MenuEvent MenuEventInfo { get; set; }
        private Form _form;

        protected virtual bool IsSystemForm { get; set; } = false;

        public XmlDocument FormXml
        {
            get
            {
                if (_formXml == null)
                {
                    try
                    {
                        string xmlFile = Path.Combine(ViewsFolder, FormID + ".srf");
                        if (File.Exists(xmlFile))
                        {
                            _formXml = new XmlDocument();
                            _formXml.Load(xmlFile);
                        }
                        else
                        {
                            string xmlString = SBOApp.ViewsResourceManager.GetString(FormID);
                            if (!String.IsNullOrEmpty(xmlString))
                            {
                                _formXml = new XmlDocument();
                                _formXml.LoadXml(xmlString);
                            }
                        }
                    }
                    catch (System.Resources.MissingManifestResourceException resourceEx)
                    {
                        //throw resourceEx;
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                }

                return _formXml;
            }
        }

        /// <summary>
        /// Pasta aonde esta o arquivo Srf (Caso nao utilize o srf no tipo StringReader)
        /// </summary>
        public virtual string ViewsFolder { get { return _viewsFolder ?? "Views"; } set { _viewsFolder = value; } }

        /// <summary>
        /// ID unico do form
        /// </summary>
        public virtual string FormID
        {
            get
            {
                string className;

                FormAttribute formAttr = typeof(BaseForm).GetCustomAttributes(typeof(FormAttribute)).FirstOrDefault() as FormAttribute;
                if (formAttr != null)
                {
                    return formAttr.FormId;
                }

                if (!_formId.TryGetValue(GetType(), out className))
                {
                    className = GetType().Name;
                    _formId[GetType()] = className;
                }

                return className;
            }

            set { _formId[GetType()] = value; }
        }

        /// <summary>
        /// Count único do form
        /// </summary>
        public virtual int FormCount
        {
            get
            {
                int count;

                if (!_formCount.TryGetValue(GetType(), out count))
                {
                    count = 0;
                }

                return count;
            }

            set { _formCount[GetType()] = value; }
        }

        /// <summary>
        /// EditText de controle da navegacao
        /// </summary>
        public string BrowseBy { get; set; }

        protected Form Form
        {
            get
            {
                if (_form == null)
                {
                    string formId = null;

                    if (ItemEventInfo != null) formId = ItemEventInfo.FormUID;
                    if (BusinessObjectInfo != null) formId = BusinessObjectInfo.FormUID;
                    if (ContextMenuInfo != null) formId = ContextMenuInfo.FormUID;

                    if (formId == null)
                        throw new Exception("Para instanciar o form é necessário estar em um formulário.");

                    _form = SBOApp.Application.Forms.Item(formId);
                }

                return _form;
            }
        }

        #endregion Properties

        #region Methods

        static BaseForm()
        {
            _formId = new Dictionary<Type, string>();
            _formCount = new Dictionary<Type, int>();
        }

        public virtual void Freeze(bool freeze)
        {
            //if (ItemEventInfo.EventType != BoEventTypes.et_FORM_UNLOAD)
            //	form = SBOApp.Application.Forms.GetFormByTypeAndCount(ItemEventInfo.FormType, ItemEventInfo.FormTypeCount);

            //if (form != null)
            //	form.Freeze(freeze);
            Form.Freeze(freeze);
        }

        #endregion Methods

        #region Events

        public virtual object Show()
        {
            if (FormXml != null)
            {
                _form = FormController.GenerateForm(FormXml, FormID, FormCount);
            }

            if (!String.IsNullOrEmpty(BrowseBy))
            {
                _form.DataBrowser.BrowseBy = BrowseBy;
            }

            return _form;
        }

        public virtual object Show(string viewsPath)
        {
            _form = FormController.GenerateForm(viewsPath, FormCount);
            return _form;
        }

        public virtual object Show(string[] args)
        {
            throw new NotImplementedException();
        }

        public virtual bool ItemEvent()
        {
            if (IsSystemForm)
            {
                if (ItemEventInfo.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    if (!ItemEventInfo.BeforeAction)
                    {
                        var xmlDocument = FormXml;
                        if (FormXml != null)
                        {
                            var form = xmlDocument.SelectSingleNode("/Application/forms/action[@type='update']/form");
                            var formUID = form.Attributes["uid"];

                            form.Attributes.RemoveAll();
                            formUID.Value = Form.UniqueID;
                            form.Attributes.Append(formUID);

                            var innerXml = xmlDocument.InnerXml;
                            SBOApp.Application.LoadBatchActions(ref innerXml);
                        }
                    }
                }
            }

            return true;
        }

        public virtual bool FormDataEvent()
        {
            return true;
        }

        public virtual bool AppEvent()
        {
            return true;
        }

        public virtual bool MenuEvent()
        {
            // Se o Form for referênciado diretamente do Menu
            if (MenuEventInfo.BeforeAction && MenuEventInfo.MenuUID == GetType().Name.Substring(1))
            {
                // Criar o formulário
                Show();

                return true;
            }
            return true;
        }

        public virtual bool PrintEvent()
        {
            return true;
        }

        public virtual bool ProgressBarEvent()
        {
            return true;
        }

        public virtual bool ReportDataEvent()
        {
            return true;
        }

        public virtual bool RightClickEvent()
        {
            return true;
        }

        public virtual bool StatusBarEvent()
        {
            return true;
        }

        public virtual bool MenuFind()
        {
            return true;
        }

        public virtual bool MenuDuplicate()
        {
            return true;
        }

        public virtual bool MenuRemove()
        {
            return true;
        }

        public virtual bool MenuAddRow()
        {
            return true;
        }

        public virtual bool MenuRemoveRow()
        {
            return true;
        }

        public virtual bool MenuAdd()
        {
            return true;
        }

        public virtual bool MenuCancel()
        {
            return true;
        }

        public virtual bool MenuFirstRecord()
        {
            return true;
        }

        public virtual bool MenuPreviousRecord()
        {
            return true;
        }

        public virtual bool MenuNextRecord()
        {
            return true;
        }

        public virtual bool MenuLastRecord()
        {
            return true;
        }

        #endregion Events
    }
}