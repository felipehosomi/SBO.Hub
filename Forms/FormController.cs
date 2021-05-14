using SAPbouiCOM;
using SBO.Hub.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace SBO.Hub.Controllers
{
    public class FormController
    {
        /// <summary>
        /// Cria o objeto do tipo Forms.IForm
        /// </summary>
        /// <typeparam name="T">Tipo do evento</typeparam>
        /// <param name="form">Tipo do Form</param>
        /// <param name="evento">Objeto do evento</param>
        /// <returns>Objeto Forms.IForm</returns>
        public static Forms.IForm CreateForm<T>(Type form, ref T evento)
        {
            ConstructorInfo constructor = form.GetConstructor(new Type[] { typeof(T) });
            Forms.IForm newForm = null;

            try
            {
                newForm = (Forms.IForm)constructor.Invoke(new object[] { evento });
            }
            catch (NullReferenceException) { }
            catch (Exception ex) { throw ex; }

            return newForm;
        }

        /// <summary>
        /// Retorna o Tipo do objeto de acordo com o ID
        /// </summary>
        /// <param name="id">Id do Form</param>
        /// <returns>Tipo do Form</returns>
        public static Type GetFormType(String id)
        {
            var assembly = SBOApp.ViewsAssembly;

            var type = assembly.GetTypes().SingleOrDefault(t => t.Name == id || t.Name == "f" + id);
            if (type == null)
            {
                List<Type> typeList = assembly.GetTypes().ToList();
                foreach (var item in typeList)
                {
                    foreach (Attribute attribute in item.GetCustomAttributes(true))
                    {
                        FormAttribute formAttribute = attribute as FormAttribute;
                        if (formAttribute != null)
                        {
                            if (formAttribute.FormId.ToString() == id)
                            {
                                return item;
                            }
                        }
                    }
                }
            }
            return type;

        }

        /// <summary>
        /// Atualiza os dados do form
        /// Utilizado em forms do B1 em que o método Form.Update() não funciona
        /// </summary>
        /// <param name="formType">Id do form</param>
        /// <returns>Executado com sucesso</returns>
        public static bool UpdateForm(int formType)
        {
            // A atualização do form é feita navegando para o próximo registro e voltando para o original
            // Não foi encontrada uma maneira mais elegante para atualizar um form do B1
            try
            {
                // Busca o form
                Form frmUpdateForm = SBOApp.Application.Forms.GetFormByTypeAndCount(formType, 1);
                // Seta o foco no form
                frmUpdateForm.Select();

                frmUpdateForm.Freeze(true);

                // Navega para o próximo item
                MenuItem menu = SBOApp.Application.Menus.Item("1288");
                menu.Enabled = true;
                menu.Activate();

                // Volta para o item original
                menu = SBOApp.Application.Menus.Item("1289");
                menu.Enabled = true;
                menu.Activate();
                frmUpdateForm.Freeze(false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void OpenForm(string srfFileName)
        {
            string appPath = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

            if (!appPath.EndsWith(@"\")) appPath += @"\";
            if (!srfFileName.EndsWith(".srf"))
                srfFileName += ".srf";

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(String.Format(@"{0}Views\{0}", appPath, srfFileName));
            string xml = xmlDoc.InnerXml;
            SBOApp.Application.LoadBatchActions(ref xml);
        }

        public static Form GenerateForm(string srfPath, int formCount)
        {
            var appPath = Environment.CurrentDirectory;
            var fullPath = Path.Combine(appPath, srfPath);
            string fileName = Path.GetFileNameWithoutExtension(fullPath);

            var xmlDoc = new XmlDocument();
            var creationPackage = (FormCreationParams)SBOApp.Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            xmlDoc.Load(fullPath);

            creationPackage.XmlData = xmlDoc.InnerXml;
            creationPackage.UniqueID = $"{fileName}_{formCount}";

            return SBOApp.Application.Forms.AddEx(creationPackage);
        }

        public static Form GenerateForm(string fileName, int formCount, string filePath = "Views")
        {
            var appPath = Environment.CurrentDirectory;
            var fullPath = Path.Combine(appPath, filePath, fileName);

            var xmlDoc = new XmlDocument();
            var creationPackage = (FormCreationParams)SBOApp.Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            xmlDoc.Load(fullPath);

            creationPackage.XmlData = xmlDoc.InnerXml;
            creationPackage.UniqueID = $"{Path.GetFileNameWithoutExtension(fullPath)}_{formCount}";

            return SBOApp.Application.Forms.AddEx(creationPackage);
        }

        public static Form GenerateForm(XmlDocument xmlDoc, string formID, int formCount)
        {
            var creationPackage =
                (FormCreationParams)SBOApp.Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);

            creationPackage.XmlData = xmlDoc.InnerXml;
            creationPackage.UniqueID = String.Format("{0}_{1}", formID, formCount);

            return SBOApp.Application.Forms.AddEx(creationPackage);
        }

        public static Form GenerateForm(StringReader srfFile, string formID, int formCount)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(srfFile);

            return GenerateForm(xmlDoc, formID, formCount);
        }

        public static void AddDataTable(Form form, string dataTableId)
        {
            for (int i = 0; i < form.DataSources.DataTables.Count; i++)
            {
                if (form.DataSources.DataTables.Item(i).UniqueID == dataTableId)
                {
                    return;
                }
            }
            form.DataSources.DataTables.Add(dataTableId);
        }

        public static Item CreateUDF(Form form, BoFormItemTypes fieldType, string fieldName, string tableName, string fieldDesription, int top, int left, int height, int width, int toPane, int fromPane, string fieldToLink, string fieldToGroupWith)
        {
            try
            {
                return CreateUDF(form, fieldType, fieldName, tableName, fieldDesription, top, left, height, width, toPane, fromPane, fieldToLink, fieldToGroupWith, fieldName);
            }

            catch (Exception ex)
            {
                SBOApp.Application.SetStatusBarMessage("Falha ao criar campos customizados no formulário: " + ex.Message);
                return null;
            }
        }


        /// <summary>
        /// cria itens de formulário de acordo com os paramêtros
        /// </summary>
        /// <param name="oForm">SAP Form</param>
        /// <param name="fieldType">Tipo de Campo (conforme padrão SAP)</param>
        /// <param name="fieldName">Nome do Campo no BD</param>
        /// <param name="tableName">Tabela do BD (DbDataSource)</param>
        /// <param name="fieldDesription">Descrição ou Caption, quando aplicável</param>
        /// <param name="top">Posição: Eixo Y</param>
        /// <param name="left">Posição: Eixo X</param>
        /// <param name="height">Tamanho: Altura</param>
        /// <param name="width">Tamanho: Comprimento</param>
        /// <param name="toPane">Painel: de</param>
        /// <param name="FromPane">Painel: até</param>
        /// <param name="FieldToLink">Item do Formulário para likar. Importante para manter reposicionamento ao redimensionar o Form</param>
        /// <param name="FieldToGroupWith">Item para agrupar a. Importate para Option Button, no qual somente uma opção do grupo pode estar selecionada ao mesmo tempo</param>
        /// <param name="formItemName">Caso queira que o FormItem não tenha o mesmo nome do parâmetro FieldName, preencher esse. Importante para itens a serem criados para campos com nome muito extensos. Pois o formItem não comporta nomes tão grandes.</param>
        /// <returns></returns>
        public static Item CreateUDF(Form form, BoFormItemTypes fieldType, string fieldName, string tableName, string fieldDesription, int top, int left, int height, int width, int toPane, int FromPane, string FieldToLink, string FieldToGroupWith, string formItemName)
        {
            Item newItem;
            StaticText staticText;
            EditText editText;
            ComboBox comboBox;
            Folder folder;
            Button button;
            Grid grid;
            Matrix matrix;
            OptionBtn optionBtn;

            try
            {
                newItem = form.Items.Add(formItemName, fieldType);
                newItem.Top = top;
                newItem.Left = left;
                newItem.Width = width;
                newItem.Height = height;
                newItem.FromPane = FromPane;
                newItem.ToPane = toPane;
                newItem.DisplayDesc = true;
                newItem.Description = fieldDesription;
                newItem.LinkTo = FieldToLink;

                switch (fieldType)
                {
                    case BoFormItemTypes.it_STATIC:
                        staticText = (StaticText)form.Items.Item(formItemName).Specific;
                        staticText.Caption = fieldDesription;
                        break;

                    case BoFormItemTypes.it_EDIT:
                        editText = (EditText)form.Items.Item(formItemName).Specific;
                        if (!string.IsNullOrEmpty(fieldName))
                            editText.DataBind.SetBound(true, tableName, fieldName);
                        break;

                    case BoFormItemTypes.it_EXTEDIT:
                        editText = (EditText)form.Items.Item(formItemName).Specific;
                        if (!string.IsNullOrEmpty(fieldName))
                            editText.DataBind.SetBound(true, tableName, fieldName);
                        break;

                    case BoFormItemTypes.it_COMBO_BOX:
                        comboBox = (ComboBox)form.Items.Item(formItemName).Specific;
                        if (!string.IsNullOrEmpty(fieldName))
                            comboBox.DataBind.SetBound(true, tableName, fieldName);
                        break;

                    case BoFormItemTypes.it_FOLDER:
                        folder = (Folder)form.Items.Item(formItemName).Specific;
                        folder.GroupWith(FieldToGroupWith);
                        folder.Caption = fieldDesription;
                        newItem.Top = form.Items.Item(FieldToGroupWith).Top;
                        newItem.Left = form.Items.Item(FieldToGroupWith).Left + form.Items.Item(FieldToGroupWith).Width;
                        break;

                    case BoFormItemTypes.it_BUTTON:
                        button = (Button)form.Items.Item(formItemName).Specific;
                        button.Caption = fieldDesription;
                        break;

                    case BoFormItemTypes.it_GRID:
                        grid = (Grid)form.Items.Item(formItemName).Specific;
                        break;

                    case BoFormItemTypes.it_MATRIX:
                        matrix = (Matrix)form.Items.Item(formItemName).Specific;
                        break;

                    case BoFormItemTypes.it_OPTION_BUTTON:
                        optionBtn = (OptionBtn)form.Items.Item(formItemName).Specific;
                        if (!string.IsNullOrEmpty(fieldName))
                            optionBtn.DataBind.SetBound(true, tableName, fieldName);
                        optionBtn.Caption = fieldDesription;


                        if (FieldToGroupWith != null && FieldToGroupWith.Length > 0)
                            optionBtn.GroupWith(FieldToGroupWith);

                        break;
                    default:
                        throw new Exception("A função não comporta o tipo de campo inserido! Efetuar inserção do campo manualmente ao formulário.");

                }

                return newItem;

            }
            catch (Exception ex)
            {
                SBOApp.Application.SetStatusBarMessage("Falha ao criar campos customizados no formulário: " + ex.Message);
                return null;
            }
        }
    }
}