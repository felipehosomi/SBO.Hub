using System;

namespace SBO.Hub.Forms
{
    public class BaseMenu : SBO.Hub.Forms.IForm
    {
        public SAPbouiCOM.MenuEvent menuEvent;

        public virtual Boolean MenuEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean AppEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean FormDataEvent()
        {
            throw new NotImplementedException();
        }

        public virtual void Freeze(Boolean freeze)
        {
            throw new NotImplementedException();
        }

        public virtual Boolean ItemEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean PrintEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean ProgressBarEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean ReportDataEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Boolean RightClickEvent()
        {
            throw new NotImplementedException();
        }

        public virtual Object Show()
        {
            throw new NotImplementedException();
        }

        public virtual Object Show(string srfPath)
        {
            throw new NotImplementedException();
        }

        public virtual Object Show(String[] args)
        {
            throw new NotImplementedException();
        }

        public virtual Boolean StatusBarEvent()
        {
            throw new NotImplementedException();
        }

        public bool MenuFind()
        {
            throw new NotImplementedException();
        }

        public bool MenuAddRow()
        {
            throw new NotImplementedException();
        }

        public bool MenuRemoveRow()
        {
            throw new NotImplementedException();
        }

        public bool MenuAdd()
        {
            throw new NotImplementedException();
        }

        public bool MenuCancel()
        {
            throw new NotImplementedException();
        }

        public bool MenuDuplicate()
        {
            throw new NotImplementedException();
        }

        public bool MenuRemove()
        {
            throw new NotImplementedException();
        }

        public bool MenuFirstRecord()
        {
            throw new NotImplementedException();
        }

        public bool MenuPreviousRecord()
        {
            throw new NotImplementedException();
        }

        public bool MenuNextRecord()
        {
            throw new NotImplementedException();
        }

        public bool MenuLastRecord()
        {
            throw new NotImplementedException();
        }
    }
}