using SAPbouiCOM;

namespace SBO.Hub.Models
{
    public class ItemEventModel
    {

        public bool ActionSuccess { get; set; }

        public bool BeforeAction { get; set; }

        public int CharPressed { get; set; }

        public string ColUID { get; set; }

        public BoEventTypes EventType { get; set; }

        public int FormMode { get; set; }

        public int FormTypeCount { get; set; }

        public string FormTypeEx { get; set; }

        public string FormUID { get; set; }

        public bool InnerEvent { get; set; }

        public bool ItemChanged { get; set; }

        public string ItemUID { get; set; }

        public BoModifiersEnum Modifiers { get; set; }

        public int PopUpIndicator { get; set; }

        public int Row { get; set; }

        public static void Update(ItemEventModel result, SAPbouiCOM.IItemEvent itemEvent)
        {
            result.ActionSuccess = itemEvent.ActionSuccess;
            result.BeforeAction = itemEvent.BeforeAction;
            result.CharPressed = itemEvent.CharPressed;
            result.ColUID = itemEvent.ColUID;
            result.EventType = itemEvent.EventType;
            result.FormMode = itemEvent.FormMode;
            result.FormTypeCount = itemEvent.FormTypeCount;
            result.FormTypeEx = itemEvent.FormTypeEx;
            result.FormUID = itemEvent.FormUID;
            result.InnerEvent = itemEvent.InnerEvent;
            result.ItemChanged = itemEvent.ItemChanged;
            result.ItemUID = itemEvent.ItemUID;
            result.Modifiers = itemEvent.Modifiers;
            result.PopUpIndicator = itemEvent.PopUpIndicator;
            result.Row = itemEvent.Row;
        }
    }
}
