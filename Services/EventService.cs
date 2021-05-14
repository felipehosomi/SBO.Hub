﻿using SAPbouiCOM;
using SBO.Hub.Controllers;
using SBO.Hub.Forms;
using System;
using System.Collections.Generic;

namespace SBO.Hub.Services
{
    public class EventService
    {
        #region Variables

        public static List<string> MenuEvents = new List<string>() { "1281", "1282", "1283", "1284", "1287", "1288", "1289", "1290", "1291", "1292", "1293" };

        public static Int32 _iItemEventRow = -1;
        public static String _sItemEventCol = String.Empty;
        public static String _sFormDataEventObjectKey = String.Empty;

        #endregion Variables

        #region Events

        public static void AppEvent(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    System.Environment.Exit(0);
                    break;
            }
        }

        public static void FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
            // The event provides the unique ID (BusinessObjectInfo.ObjectKey) of the modified business object.
            _sFormDataEventObjectKey = BusinessObjectInfo.ObjectKey;

            // Executa o método FormDataEvent do formulário em que ocorreu o evento
            BubbleEvent = ExecuteEvent<BusinessObjectInfo>(BusinessObjectInfo.FormTypeEx, BusinessObjectInfo, "FormDataEvent", true);
        }

        public static void ItemEvent(String FormUID, ref ItemEvent itemEvent, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
            if (itemEvent.EventType == BoEventTypes.et_FORM_UNLOAD && !itemEvent.BeforeAction)
            {
                return;
            }

            // Obtém a linha e a coluna em que o evento está ocorrendo, caso se trate de um grid ou matrix
            _iItemEventRow = itemEvent.Row;
            _sItemEventCol = itemEvent.ColUID;

            // Executa o método ItemEvent do formulário em que ocorreu o evento
            BubbleEvent = ExecuteEvent<ItemEvent>(itemEvent.FormTypeEx, itemEvent, "ItemEvent", true);
        }

        public static void MenuEvent(ref MenuEvent menuEvent, out Boolean BubbleEvent)
        {
            try
            {
                Type tType = null;
                if (MenuEvents.Contains(menuEvent.MenuUID))
                {
                    tType = FormController.GetFormType(SBOApp.Application.Forms.ActiveForm.TypeEx);
                    BaseForm form = (BaseForm)Activator.CreateInstance(tType, menuEvent);
                    switch (menuEvent.MenuUID)
                    {
                        case "1281":
                            BubbleEvent = form.MenuFind();
                            break;
                        case "1282":
                            BubbleEvent = form.MenuAdd();
                            break;
                        case "1283":
                            BubbleEvent = form.MenuRemove();
                            break;
                        case "1284":
                            BubbleEvent = form.MenuCancel();
                            break;
                        case "1287":
                            BubbleEvent = form.MenuDuplicate();
                            break;
                        case "1288":
                            BubbleEvent = form.MenuNextRecord();
                            break;
                        case "1289":
                            BubbleEvent = form.MenuPreviousRecord();
                            break;
                        case "1290":
                            BubbleEvent = form.MenuFirstRecord();
                            break;
                        case "1291":
                            BubbleEvent = form.MenuLastRecord();
                            break;
                        case "1292":
                            BubbleEvent = form.MenuAddRow();
                            break;
                        case "1293":
                            BubbleEvent = form.MenuRemoveRow();
                            break;

                        default:
                            BubbleEvent = ExecuteEvent<MenuEvent>(menuEvent.MenuUID, menuEvent, "MenuEvent", false);
                            break;
                    }
                }
                else
                {
                    if (!menuEvent.BeforeAction)
                    {
                        tType = FormController.GetFormType(menuEvent.MenuUID);
                        if (tType == null)
                        {
                            BubbleEvent = ExecuteEvent<MenuEvent>(menuEvent.MenuUID, menuEvent, "MenuEvent", false);
                        }
                        else
                        {
                            try
                            {
                                ((BaseForm)Activator.CreateInstance(tType)).Show();
                            }
                            catch (Exception ex)
                            {
                                BubbleEvent = ExecuteEvent<MenuEvent>(menuEvent.MenuUID, menuEvent, "MenuEvent", false);
                            }
                            BubbleEvent = true;
                        }
                    }
                    else
                    {
                        BubbleEvent = ExecuteEvent<MenuEvent>(menuEvent.MenuUID, menuEvent, "MenuEvent", false);
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = true;
            }
        }

        public static void RightClickEvent(ref ContextMenuInfo contextMenuInfo, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
            // Executa o método RightClickEvent do formulário em que ocorreu o evento
            ExecuteEvent<ContextMenuInfo>(SBOApp.Application.Forms.Item(contextMenuInfo.FormUID.ToString()).TypeEx,
                                                        contextMenuInfo,
                                                        "RightClickEvent",
                                                        false);
        }

        #endregion Events

        #region ExecuteEvent

        /// <summary>
        /// Executa evento
        /// </summary>
        /// <typeparam name="T">Tipo do evento</typeparam>
        /// <param name="formID">ID do form</param>
        /// <param name="eventInfo">Objeto do evento</param>
        /// <param name="eventName">Nome do evento a ser executado</param>
        /// <param name="finishTransactionYN">Finaliza transação em caso de erro</param>
        /// <returns>Evento foi executado?</returns>
        public static Boolean ExecuteEvent<T>(String formID, T eventInfo, String eventName, Boolean finishTransactionYN)
        {
            Forms.IForm oForm = null;

            try
            {
                Type tType = FormController.GetFormType(formID);
                // Verifica se a variável tType contém algum valor, caso não, o método é interrompido
                if (tType == null) return true;

                // Instancia o objeto referente ao menu que foi selecionado
                oForm = FormController.CreateForm<T>(tType, ref eventInfo);
                // Verifica se a variável oForm contém algum valor, caso não, o método é interrompido
                //if (oForm == null) return false;
                if (oForm == null) return true;

                return (Boolean)oForm.GetType().GetMethod(eventName).Invoke(oForm, null);
            }
            catch (Exception ex)
            {
                SBOApp.Application.SetStatusBarMessage(ex.Message);
                return false;
            }
        }

        #endregion ExecuteEvent

        #region NonImplementedEvents

        #region PrintEvent

        public static void PrintEvent(ref PrintEventInfo eventInfo, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion PrintEvent

        #region ProgressBarEvent

        public static void ProgressBarEvent(ref ProgressBarEvent progressBarEvent, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion ProgressBarEvent

        #region ReportDataEvent

        public static void ReportDataEvent(ref ReportDataInfo reportDataInfo, out Boolean BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion ReportDataEvent

        #region StatusBarEvent

        public static void StatusBarEvent(String Text, BoStatusBarMessageType MessageType)
        {
            if (Text.Contains("UI_API -7780") && !String.IsNullOrEmpty(SBOApp.StatusBarMessage))
            {
                SBOApp.Application.StatusBar.SetText(SBOApp.StatusBarMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion StatusBarEvent

        #endregion NonImplementedEvents
    }
}