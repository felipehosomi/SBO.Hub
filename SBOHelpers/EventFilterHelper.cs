using SAPbouiCOM;
using System;

namespace SBO.Hub.Helpers
{
    public class EventFilterHelper
    {
        private static EventFilters EventFilters;

        private static EventFilters _disableEventFilters;

        private static EventFilters DisableEventFilters
        {
            get
            {
                if (_disableEventFilters == null)
                {
                    _disableEventFilters = new EventFilters();
                    _disableEventFilters.Reset();
                }

                return _disableEventFilters;
            }
        }

        public static void CreateDefaultEvents()
        {
            EventFilters = new EventFilters();

            // Sempre adicionar MENU CLICK, se não os menus não abrem
            EventFilter filter = EventFilters.Add(BoEventTypes.et_MENU_CLICK);
            SBOApp.Application.SetFilter(EventFilters);
        }

        public static void SetDefaultEvents()
        {
            EventFilters = new EventFilters();

            // Sempre adicionar MENU CLICK, se não os menus não abrem
            EventFilter filter = EventFilters.Add(BoEventTypes.et_MENU_CLICK);

            // Adiciona eventos que não impactam na performance
            filter = EventFilters.Add(BoEventTypes.et_CLICK);
            filter = EventFilters.Add(BoEventTypes.et_CHOOSE_FROM_LIST);
            filter = EventFilters.Add(BoEventTypes.et_FORM_CLOSE);
            filter = EventFilters.Add(BoEventTypes.et_COMBO_SELECT);
            filter = EventFilters.Add(BoEventTypes.et_RIGHT_CLICK);
            filter = EventFilters.Add(BoEventTypes.et_DOUBLE_CLICK);

            filter = EventFilters.Add(BoEventTypes.et_FORM_RESIZE);

            // Data Form Events
            filter = EventFilters.Add(BoEventTypes.et_FORM_DATA_ADD);
            filter = EventFilters.Add(BoEventTypes.et_FORM_DATA_UPDATE);
            filter = EventFilters.Add(BoEventTypes.et_FORM_DATA_LOAD);
            filter = EventFilters.Add(BoEventTypes.et_FORM_DATA_DELETE);

            // Adiciona FormLoad para poder setar o filtro nos system forms
            filter = EventFilters.Add(BoEventTypes.et_FORM_LOAD);
            filter = EventFilters.Add(BoEventTypes.et_PRINT_LAYOUT_KEY);

            SBOApp.Application.SetFilter(EventFilters);
        }

        /// <summary>
        /// Aciona evento no Form
        /// </summary>
        /// <param name="formId">Id do Form - Exemplos: 150 / 2000002001</param>
        /// <param name="eventType">Tipo do evento</param>
        public static void SetFormEvent(string formId, BoEventTypes eventType)
        {
            try
            {
                if (EventFilters == null)
                {
                    CreateDefaultEvents();
                }

                EventFilter filter;
                // Busca o evento na lista de eventos
                for (int i = 0; i < EventFilters.Count; i++)
                {
                    filter = EventFilters.Item(i);
                    if (filter.EventType == eventType)
                    {
                        try
                        {
                            filter.AddEx(formId);
                        }
                        catch (Exception e) { }

                        return;
                    }
                }

                // Se não encontrar o evento, adiciona
                filter = EventFilters.Add(eventType);
                if (!String.IsNullOrEmpty(formId))
                {
                    filter.AddEx(formId);
                }
            }
            catch (Exception e)
            {
                SBOApp.Application.SetStatusBarMessage("Erro geral no EventFilter: " + e.Message);
            }
        }

        /// <summary>
        /// Desabilita todos os eventos
        /// </summary>
        public static void DisableEvents()
        {
            SBOApp.Application.SetFilter(DisableEventFilters);
        }

        public static void EnableEvents()
        {
            SBOApp.Application.SetFilter(EventFilters);
        }
    }
}