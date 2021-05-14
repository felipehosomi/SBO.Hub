using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using System.Reflection;
using SBO.Hub.Properties;
using System.Globalization;
using SBO.Hub.Attributes;

namespace SBO.Hub.UI
{
    public class ValidationController
    {
        public string LastMessageString { get; set; }

        /// <summary>
        /// Preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="form">Formulário que contém os campos a serem validados</param>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public T FillModel<T>(Form form)
        {
            T model = Activator.CreateInstance<T>();
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    this.ValidateField(form, property, out value);

                    // Se valor do campo for vazio, vai para o próximo campo
                    if (String.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    // Seta valor no model
                    if (property.PropertyType == typeof(string))
                    {
                        property.SetValue(model, value, null);
                    }
                    else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>))
                    {
                        property.SetValue(model, Convert.ToInt32(value), null);
                    }
                    else if (property.PropertyType == typeof(Int16) || property.PropertyType == typeof(Nullable<Int16>))
                    {
                        property.SetValue(model, Convert.ToInt16(value), null);
                    }
                    else if (property.PropertyType == typeof(Int64) || property.PropertyType == typeof(Nullable<Int64>))
                    {
                        property.SetValue(model, Convert.ToInt64(value), null);
                    }
                    else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(Nullable<decimal>))
                    {
                        if (value.StartsWith("."))
                        {
                            value = "0" + value;
                        }
                        value = value.Replace(".", ",");
                        property.SetValue(model, Convert.ToDecimal(value), null);
                    }
                    else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<double>))
                    {
                        if (value.StartsWith("."))
                        {
                            value = "0" + value;
                        }
                        value = value.Replace(".", ",");
                        property.SetValue(model, Convert.ToDouble(value), null);
                    }
                    else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                    {
                        DateTime dateTimeValue;
                        if (DateTime.TryParseExact(value, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTimeValue))
                        {
                            property.SetValue(model, dateTimeValue, null);
                        }
                        else if (DateTime.TryParse(value, out dateTimeValue))
                        {
                            property.SetValue(model, dateTimeValue, null);
                        }
                    }
                    else
                    {
                        property.SetValue(model, value, null);
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(String.Format("Erro ao preencher propriedade {0} com o valor {1}: {2}", property.Name, value, e.Message));
                }
            }
            return model;
        }

        /// <summary>
        /// Valida se campos obrigatórios estão preenchidos e preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="form">Formulário que contém os campos a serem validados</param>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public bool ValidateAndFillModel<T>(Form form, ref T model)
        {
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    // Se houver erro em algum campo, retorna falso
                    if (!this.ValidateField(form, property, out value))
                    {
                        return false;
                    }

                    // Se valor do campo for vazio, vai para o próximo campo
                    if (String.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    // Seta valor no model
                    switch (Type.GetTypeCode(property.PropertyType))
                    {
                        case TypeCode.DateTime:
                            DateTime dateTimeValue;
                            if (DateTime.TryParseExact(value, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTimeValue))
                            {
                                property.SetValue(model, dateTimeValue, null);
                            }
                            else if (DateTime.TryParse(value, out dateTimeValue))
                            {
                                property.SetValue(model, dateTimeValue, null);
                            }
                            break;
                        case TypeCode.Decimal:
                            property.SetValue(model, Convert.ToDecimal(value), null);
                            break;
                        case TypeCode.Double:
                            if (value.StartsWith("."))
                            {
                                value = "0" + value;
                            }
                            value = value.Replace(".", ",");

                            property.SetValue(model, Convert.ToDouble(value), null);
                            break;
                        case TypeCode.Int16:
                            property.SetValue(model, Convert.ToInt16(value), null);
                            break;
                        case TypeCode.Int32:
                            property.SetValue(model, Convert.ToInt32(value), null);
                            break;
                        case TypeCode.Int64:
                            property.SetValue(model, Convert.ToInt64(value), null);
                            break;
                        case TypeCode.String:
                            property.SetValue(model, value, null);
                            break;
                        default:
                            property.SetValue(model, value, null);
                            break;
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(String.Format("Erro ao preencher propriedade {0} com o valor {1}: {2}", property.Name, value, e.Message));
                }
            }
            return true;
        }

        public bool ValidateFields(Form form, Type modelType)
        {
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                string value;
                if (!this.ValidateField(form, property, out value))
                {
                    return false;
                }
            }
            return true;
        }

        private bool ValidateField(Form form, PropertyInfo property, out string value)
        {
            value = String.Empty;
            foreach (Attribute attribute in property.GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
                    if (!hubModel.MandatoryYN)
                    {
                        continue;
                    }

                    // Se não existir o nome do campo na tela, vai para o próximo
                    if (String.IsNullOrEmpty(hubModel.UIFieldName))
                    {
                        return true;
                    }

                    Item item = null;
                    // Busca o item
                    try
                    {
                        item = form.Items.Item(hubModel.UIFieldName);
                    }
                    catch (Exception e)
                    {
                        LastMessageString = String.Format(CommonStrings.GetItemError, hubModel.UIFieldName, e.Message);
                        SBOApp.Application.SetStatusBarMessage(LastMessageString);
                        return false;
                    }

                    // Busca o conteúdo do item
                    switch (item.Type)
                    {
                        case BoFormItemTypes.it_CHECK_BOX:
                            value = ((CheckBox)item.Specific).Checked ? ((CheckBox)item.Specific).ValOn : ((CheckBox)item.Specific).ValOff;
                            break;
                        case BoFormItemTypes.it_COMBO_BOX:
                            value = ((ComboBox)item.Specific).Value;
                            break;
                        case BoFormItemTypes.it_EDIT:
                            value = ((EditText)item.Specific).Value;
                            break;
                    }
                    if (String.IsNullOrEmpty(value))
                    {
                        // Seta mensagem na tela caso seja obrigatório e estiver vazio
                        LastMessageString = String.Format(CommonStrings.FieldMustBeFilled, hubModel.Description);
                        SBOApp.Application.SetStatusBarMessage(LastMessageString);
                        item.Click();
                        return false;

                    }
                }
            }
            return true;
        }
    }
}
