using SAPbouiCOM;
using SBO.Hub.Attributes;
using SBO.Hub.Properties;
using System;
using System.Globalization;
using System.Reflection;

namespace SBO.Hub.UI
{
    public static class FormExtensions
    {
        #region ClearFields
        /// <summary>
        /// Limpa os campos da tela
        /// </summary>
        public static void ClearFields(this Form form)
        {
            for (int i = 0; i < form.Items.Count; i++)
            {
                switch (form.Items.Item(i).Type)
                {
                    case BoFormItemTypes.it_EXTEDIT:
                    case BoFormItemTypes.it_EDIT:
                        ((EditText)form.Items.Item(i).Specific).Value = String.Empty;
                        break;
                    case BoFormItemTypes.it_COMBO_BOX:
                        ((ComboBox)form.Items.Item(i).Specific).Select(0, BoSearchKey.psk_Index);
                        break;
                    case BoFormItemTypes.it_CHECK_BOX:
                        ((CheckBox)form.Items.Item(i).Specific).Checked = false;
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion ClearFields

        #region FillUIFields
        /// <summary>
        /// Carrega dados do model nos campos tela
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static T FillUIFields<T>(this Form form, T model)
        {
            HubModelAttribute hubModel;
            // Percorre as propriedades do Model
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;

                    if (hubModel != null)
                    {
                        if (String.IsNullOrEmpty(hubModel.ColumnName))
                            hubModel.ColumnName = property.Name;
                        object value = property.GetValue(model, null);

                        if (value == null)
                        {
                            value = String.Empty;
                        }

                        Item item = null;
                        // Busca o item
                        try
                        {
                            item = form.Items.Item(hubModel.UIFieldName);
                            if (!item.Enabled) // Não preenche campos desativados
                            {
                                continue;
                            }
                        }
                        catch (Exception e)
                        {
                            SBOApp.StatusBarMessage = String.Format(CommonStrings.GetItemError, hubModel.UIFieldName, e.Message);
                        }

                        // Busca o conteúdo do item
                        switch (item.Type)
                        {
                            case BoFormItemTypes.it_CHECK_BOX:
                                ((CheckBox)item.Specific).Checked = value.ToString() == ((CheckBox)item.Specific).ValOn;
                                break;

                            case BoFormItemTypes.it_COMBO_BOX:
                                ((ComboBox)item.Specific).Select(value.ToString());
                                break;

                            case BoFormItemTypes.it_EDIT:
                                ((EditText)item.Specific).Value = value.ToString();
                                break;
                        }

                    }
                    break;
                }
            }
            return model;
        }
        #endregion FillUIFields

        #region FillUserDataSources
        /// <summary>
        /// Carrega dados do model na tela preenchendo UserDataSources
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static T FillUserDataSources<T>(this Form form, T model)
        {
            HubModelAttribute hubModel;
            // Percorre as propriedades do Model
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                // Busca os Custom Attributes
                foreach (Attribute attribute in property.GetCustomAttributes(true))
                {
                    hubModel = attribute as HubModelAttribute;

                    if (hubModel != null)
                    {
                        if (String.IsNullOrEmpty(hubModel.UserDataSource))
                            hubModel.UserDataSource = property.Name;
                        object value = property.GetValue(model, null);
                        if (value == null)
                        {
                            value = String.Empty;
                        }

                        form.DataSources.UserDataSources.Item(hubModel.UserDataSource).Value = value.ToString();
                    }
                    break;
                }
            }

            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {

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
        #endregion FillUserDataSources

        #region FillModelByUI
        /// <summary>
        /// Preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static T FillModelByUI<T>(this Form form)
        {
            T model = Activator.CreateInstance<T>();
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    ValidateFieldByUI(form, property, out value);

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
        #endregion FillModelByUI

        #region ValidateAndFillModelByUI
        /// <summary>
        /// Valida se campos obrigatórios estão preenchidos e preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="form">Formulário que contém os campos a serem validados</param>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static bool ValidateAndFillModelByUI<T>(this Form form, ref T model)
        {
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    // Se houver erro em algum campo, retorna falso
                    if (!ValidateFieldByUI(form, property, out value))
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
        #endregion ValidateAndFillModelByUI

        #region ValidateFieldsByUI
        public static bool ValidateFieldsByUI(this Form form, Type modelType)
        {
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                string value;
                if (!ValidateFieldByUI(form, property, out value))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool ValidateFieldByUI(this Form form, PropertyInfo property, out string value)
        {
            value = String.Empty;
            foreach (Attribute attribute in property.GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
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
                        SBOApp.StatusBarMessage = String.Format(CommonStrings.GetItemError, hubModel.UIFieldName, e.Message);
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

                    // Seta mensagem na tela caso seja obrigatório e estiver vazio
                    if (String.IsNullOrEmpty(value) && hubModel.MandatoryYN)
                    {
                        if (String.IsNullOrEmpty(hubModel.Description))
                        {
                            hubModel.Description = property.Name;
                        }

                        SBOApp.StatusBarMessage = String.Format(CommonStrings.FieldMustBeFilled, hubModel.Description);
                        item.Click();
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion ValidateFieldsByUI

        #region ValidateFieldsByDBDataSource
        /// <summary>
        /// Valida se campos obrigatórios estão preenchidos e preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="form">Formulário que contém os campos a serem validados</param>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static bool ValidateAndFillModelByDBDataSource<T>(this Form form, ref T model)
        {
            string tableName = String.Empty;
            foreach (Attribute attribute in model.GetType().GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
                    tableName = hubModel.TableName;
                }
            }

            DBDataSource dBDataSource = null;
            try
            {

                if (String.IsNullOrEmpty(tableName))
                {
                    dBDataSource = form.DataSources.DBDataSources.Item(0);
                }
                else
                {
                    dBDataSource = form.DataSources.DBDataSources.Item(tableName);
                }
            }
            catch (Exception e)
            {
                SBOApp.StatusBarMessage = String.Format(CommonStrings.TableNameNotFound, tableName, e.Message);
                return false;
            }

            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    // Se houver erro em algum campo, retorna falso
                    if (!ValidateFieldByDBDataSource(form, dBDataSource, property, out value))
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
                            value = value.Replace(".", ",");
                            property.SetValue(model, Convert.ToDecimal(value), null);
                            break;

                        case TypeCode.Double:
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

        public static bool ValidateFieldsByDBDataSource(this Form form, Type modelType)
        {
            string tableName = String.Empty;
            foreach (Attribute attribute in modelType.GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
                    tableName = hubModel.TableName;
                }
            }

            DBDataSource dBDataSource = null;
            try
            {
                if (String.IsNullOrEmpty(tableName))
                {
                    dBDataSource = form.DataSources.DBDataSources.Item(0);
                }
                else
                {
                    dBDataSource = form.DataSources.DBDataSources.Item(tableName);
                }
            }
            catch (Exception e)
            {
                SBOApp.StatusBarMessage = String.Format(CommonStrings.TableNameNotFound, tableName, e.Message);
                return false;
            }

            foreach (PropertyInfo property in modelType.GetProperties())
            {
                string value;
                if (!ValidateFieldByDBDataSource(form, dBDataSource, property, out value))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool ValidateFieldByDBDataSource(this Form form, DBDataSource dBDataSource, PropertyInfo property, out string value)
        {
            value = String.Empty;
            foreach (Attribute attribute in property.GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
                    // Se não existir o nome do campo na tela, vai para o próximo
                    if (String.IsNullOrEmpty(hubModel.ColumnName))
                    {
                        hubModel.ColumnName = property.Name;
                    }

                    value = dBDataSource.GetValue(hubModel.ColumnName, 0).Trim();
                    // Seta mensagem na tela caso seja obrigatório e estiver vazio
                    if (String.IsNullOrEmpty(value) && hubModel.MandatoryYN)
                    {
                        if (String.IsNullOrEmpty(hubModel.Description))
                        {
                            hubModel.Description = property.Name;
                        }
                        SBOApp.StatusBarMessage = String.Format(CommonStrings.FieldMustBeFilled, hubModel.Description);
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion ValidateFieldsByDBDataSource

        /// <summary>
        /// Valida se campos obrigatórios estão preenchidos e preenche o model
        /// </summary>
        /// <typeparam name="T">Tipo do model</typeparam>
        /// <param name="form">Formulário que contém os campos a serem validados</param>
        /// <param name="model">Model preenchido</param>
        /// <returns>Validado com sucesso</returns>
        public static bool ValidateAndFillModelByUserDataSource<T>(this Form form, ref T model)
        {
            foreach (PropertyInfo property in model.GetType().GetProperties())
            {
                string value = String.Empty;
                try
                {
                    // Se houver erro em algum campo, retorna falso
                    if (!ValidateFieldByUserDataSource(form, property, out value))
                    {
                        return false;
                    }

                    // Se valor do campo for vazio, vai para o próximo campo
                    if (String.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    // Seta o valor na propriedade do model
                    if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(Nullable<DateTime>))
                    {
                        DateTime dateTimeValue;
                        if (DateTime.TryParseExact(value, "dd/MM/yyyy", System.Globalization.CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTimeValue))
                        {
                            property.SetValue(model, dateTimeValue, null);
                        }
                        else if (DateTime.TryParse(value, out dateTimeValue))
                        {
                            property.SetValue(model, dateTimeValue, null);
                        }
                    }
                    else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(Nullable<decimal>))
                    {
                        value = value.Replace(".", ",");
                        property.SetValue(model, Convert.ToDecimal(value), null);
                    }
                    else if (property.PropertyType == typeof(double) || property.PropertyType == typeof(Nullable<double>))
                    {
                        value = value.Replace(".", ",");
                        property.SetValue(model, Convert.ToDouble(value), null);
                    }
                    else if (property.PropertyType == typeof(Int16) || property.PropertyType == typeof(Nullable<Int16>))
                    {
                        property.SetValue(model, Convert.ToInt16(value), null);
                    }
                    else if (property.PropertyType == typeof(Int32) || property.PropertyType == typeof(Nullable<Int32>))
                    {
                        property.SetValue(model, Convert.ToInt32(value), null);
                    }
                    else if (property.PropertyType == typeof(Int64) || property.PropertyType == typeof(Nullable<Int64>))
                    {
                        property.SetValue(model, Convert.ToInt64(value), null);
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
            return true;
        }

        public static bool ValidateFieldsByUserDataSource(this Form form, Type modelType)
        {
            foreach (PropertyInfo property in modelType.GetProperties())
            {
                string value;
                if (!ValidateFieldByUserDataSource(form, property, out value))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool ValidateFieldByUserDataSource(this Form form, PropertyInfo property, out string value)
        {
            value = String.Empty;
            foreach (Attribute attribute in property.GetCustomAttributes(true))
            {
                HubModelAttribute hubModel = attribute as HubModelAttribute;
                if (hubModel != null)
                {
                    // Se não existir o nome do campo na tela, vai para o próximo
                    if (String.IsNullOrEmpty(hubModel.UserDataSource))
                    {
                        return true;
                    }

                    UserDataSource userDataSource = null;

                    // Busca o item
                    try
                    {
                        userDataSource = form.DataSources.UserDataSources.Item(hubModel.UserDataSource);
                    }
                    catch (Exception e)
                    {
                        SBOApp.StatusBarMessage = String.Format(CommonStrings.GetItemError, hubModel.UIFieldName, e.Message);
                        return false;
                    }

                    value = userDataSource.Value.Trim();
                    // Seta mensagem na tela caso seja obrigatório e estiver vazio
                    if (String.IsNullOrEmpty(value) && hubModel.MandatoryYN)
                    {
                        if (String.IsNullOrEmpty(hubModel.Description))
                        {
                            hubModel.Description = property.Name;
                        }
                        SBOApp.StatusBarMessage = String.Format(CommonStrings.FieldMustBeFilled, hubModel.Description);
                        return false;
                    }
                }
            }
            return true;
        }
    }
}