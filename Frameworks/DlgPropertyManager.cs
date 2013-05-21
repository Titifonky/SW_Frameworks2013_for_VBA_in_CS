using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("94E12A11-CED8-4629-9A96-63BD2DA46BD9")]
    public interface IDlgPropertyManager : IPropertyManagerPage2Handler9
    {
        Boolean Init(eSldWorks Sw, String Titre, swPropertyManagerPageOptions_e OptionsPage = 0);
        void Show();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("67CEC463-2B2A-4089-BE45-1B141724F5A9")]
    [ProgId("Frameworks.DlgPropertyManager")]
    public class DlgPropertyManager : IDlgPropertyManager
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eSldWorks _SW;
        private PropertyManagerPage2 _Page;

        #endregion

        #region "Constructeur\Destructeur"

        public DlgPropertyManager(){}

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Test l'initialisation de l'objet DlgPropertyManager.
        /// </summary>
        public Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Initialiser l'objet DlgPropertyManager.
        /// </summary>
        /// <param name="Sw"></param>
        /// <param name="Titre"></param>
        /// <param name="OptionsPage"></param>
        /// <returns></returns>
        public Boolean Init(eSldWorks Sw, String Titre, swPropertyManagerPageOptions_e OptionsPage = 0)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Sw != null) && Sw.EstInitialise)
            {
                _SW = Sw;

                if (OptionsPage == 0)
                    OptionsPage = swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage | swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;

                int longerrors = 0;
                _Page = _SW.SwSW.CreatePropertyManagerPage(Titre, (int)OptionsPage, this, ref longerrors);

                if (longerrors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
                {
                    // On valide l'initialisation
                    _EstInitialise = true;
                }

            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public void Show()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_EstInitialise)
                _Page.Show2(0);
        }

        void IPropertyManagerPage2Handler9.AfterActivation()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.AfterClose()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        int IPropertyManagerPage2Handler9.OnActiveXControlCreated(int Id, bool Status)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnButtonPress(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnCheckboxCheck(int Id, bool Checked)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnClose(int Reason)
        {
            if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel)
            {
                //Do something when the cancel button is clicked
            }
            else if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                //Do something else when the OK button is clicked
            }
        }
        void IPropertyManagerPage2Handler9.OnComboboxEditChanged(int Id, string Text)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnComboboxSelectionChanged(int Id, int Item)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnGainedFocus(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnGroupCheck(int Id, bool Checked)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnGroupExpand(int Id, bool Expanded)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnHelp()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnListboxSelectionChanged(int Id, int Item)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnLostFocus(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnNextPage()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnNumberboxChanged(int Id, double Value)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnOptionCheck(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnPopupMenuItem(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnPopupMenuItemUpdate(int Id, ref int retval)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnPreview()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnPreviousPage()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnRedo()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutCreated(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSelectionboxCalloutDestroyed(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSelectionboxFocusChanged(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSelectionboxListChanged(int Id, int Count)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSliderPositionChanged(int Id, double Value)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnSliderTrackingCompleted(int Id, double Value)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        bool IPropertyManagerPage2Handler9.OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            // This method must return true for selections to occur
            return true;
        }
        bool IPropertyManagerPage2Handler9.OnTabClicked(int Id)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnTextboxChanged(int Id, string Text)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnUndo()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnWhatsNew()
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnListboxRMBUp(int Id, int PosX, int PosY)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        int IPropertyManagerPage2Handler9.OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        void IPropertyManagerPage2Handler9.OnNumberBoxTrackingCompleted(int Id, double Value)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        #endregion
    }
}
