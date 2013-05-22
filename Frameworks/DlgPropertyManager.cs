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
        PropertyManagerPage2 SwPropertyManagerPage2 { get; }
        String Titre { get; set; }
        Boolean Epingle { get; set; }
        swPropertyManagerPageCloseReasons_e ResultatAffichage { get; }
        Boolean Init(eSldWorks Sw, swPropertyManagerPageOptions_e OptionsPage = 0);
        swPropertyManagerPageStatus_e Afficher();
        Boolean DefinirLeMessage(String Message, swPropertyManagerPageMessageVisibility Visibilite, swPropertyManagerPageMessageExpanded Deplie, String Info);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("67CEC463-2B2A-4089-BE45-1B141724F5A9")]
    [ProgId("Frameworks.DlgPropertyManager")]
    public class DlgPropertyManager : IDlgPropertyManager
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eSldWorks _SW;
        private PropertyManagerPage2 _SwPage;
        private swPropertyManagerPageCloseReasons_e _ResultatAffichage;

        #endregion

        #region "Constructeur\Destructeur"

        public DlgPropertyManager(){}

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Component2 associé.
        /// </summary>
        public PropertyManagerPage2 SwPropertyManagerPage2 { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwPage; } }

        /// <summary>
        /// Retourne ou défini le titre.
        /// </summary>
        public String Titre
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_EstInitialise)
                    return _SwPage.Title;

                return "";
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_EstInitialise)
                    _SwPage.Title = value;
            }
        }

        /// <summary>
        /// Retourne ou défini si la page est epinglée.
        /// </summary>
        public Boolean Epingle
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_EstInitialise)
                    return _SwPage.Pinned;

                return false;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_EstInitialise)
                    _SwPage.Pinned = value;
            }
        }

        /// <summary>
        /// Retourne le resultat de l'affichage de la page.
        /// </summary>
        public swPropertyManagerPageCloseReasons_e ResultatAffichage
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _ResultatAffichage;
            }
        }

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
        public Boolean Init(eSldWorks Sw, swPropertyManagerPageOptions_e OptionsPage = 0)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Sw != null) && Sw.EstInitialise)
            {
                _SW = Sw;

                if (OptionsPage == 0)
                    OptionsPage = swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage | swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;

                int longerrors = 0;
                _SwPage = _SW.SwSW.CreatePropertyManagerPage("", (int)OptionsPage, this, ref longerrors);

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

        /// <summary>
        /// Afficher la page.
        /// </summary>
        public swPropertyManagerPageStatus_e Afficher()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_EstInitialise)
                return (swPropertyManagerPageStatus_e)_SwPage.Show2(0);

            return 0;
        }

        /// <summary>
        /// Definir le message en entête de page.
        /// </summary>
        /// <param name="Message"></param>
        /// <param name="Visibilite"></param>
        /// <param name="Deplie"></param>
        /// <param name="Info"></param>
        /// <returns></returns>
        public Boolean DefinirLeMessage(String Message, swPropertyManagerPageMessageVisibility Visibilite, swPropertyManagerPageMessageExpanded Deplie, String Info)
        {
            Boolean t;
            t = _SwPage.SetMessage3(Message, (int)Visibilite, (int)Deplie, Info);
            return t;
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
            _ResultatAffichage = (swPropertyManagerPageCloseReasons_e)Reason;
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
