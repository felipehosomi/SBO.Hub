using System;

namespace SBO.Hub.Forms
{
    public interface IForm
    {
        Boolean AppEvent();

        Boolean FormDataEvent();

        void Freeze(Boolean freeze);

        Boolean ItemEvent();

        Boolean MenuEvent();

        Boolean PrintEvent();

        Boolean ProgressBarEvent();

        Boolean ReportDataEvent();

        Boolean RightClickEvent();

        Object Show();

        Object Show(string srfPath);

        Object Show(String[] args);

        Boolean StatusBarEvent();

        bool MenuFind();

        bool MenuDuplicate();

        bool MenuRemove();

        bool MenuAddRow();

        bool MenuRemoveRow();

        bool MenuAdd();

        bool MenuCancel();

        bool MenuFirstRecord();

        bool MenuPreviousRecord();

        bool MenuNextRecord();

        bool MenuLastRecord();
    }
}