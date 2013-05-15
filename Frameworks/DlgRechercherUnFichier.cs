using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("5F43729A-1E72-48F7-AAA5-0474D859E79B")]
    public interface IDlgRechercherUnFichier
    {
        String Titre { get; set; }
        Boolean AjouterExtension { get; set; }
        Boolean TestFichierExiste { get; set; }
        Boolean TestCheminExiste { get; set; }
        String ExtensionParDefaut { get; set; }
        String Filtre { get; set; }
        int IndexDuFiltre { get; set; }
        String DossierInitial { get; set; }
        Boolean RestaurerLeDossierActif { get; set; }
        Boolean AfficherAide { get; set; }
        Boolean AfficherLectureSeule { get; set; }
        Boolean ExtensionsMultiple { get; set; }

        Boolean LectureSeuleSelectionne { get; set; }

        String SelectionnerUnFichier(Boolean CheminComplet = true);
        String[] SelectionnerPlusieursFichiers(Boolean CheminComplet = true);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("736A58F3-6B78-4826-A919-82E817206398")]
    [ProgId("Frameworks.DlgRechercherUnFichier")]
    public class DlgRechercherUnFichier : IDlgRechercherUnFichier
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public DlgRechercherUnFichier() { }

        #endregion

        #region "Propriétés"

        private OpenFileDialog _Dialogue = new OpenFileDialog();

        public String Titre { get { return _Dialogue.Title; } set { _Dialogue.Title = value; } }
        public Boolean AjouterExtension { get { return _Dialogue.AddExtension; } set { _Dialogue.AddExtension = value; } }
        public Boolean TestFichierExiste { get { return _Dialogue.CheckFileExists; } set { _Dialogue.CheckFileExists = value; } }
        public Boolean TestCheminExiste { get { return _Dialogue.CheckPathExists; } set { _Dialogue.CheckPathExists = value; } }
        public String ExtensionParDefaut { get { return _Dialogue.DefaultExt; } set { _Dialogue.DefaultExt = value; } }
        public String Filtre { get { return _Dialogue.Filter; } set { _Dialogue.Filter = value; } }
        public int IndexDuFiltre { get { return _Dialogue.FilterIndex; } set { _Dialogue.FilterIndex = value; } }
        public String DossierInitial { get { return _Dialogue.InitialDirectory; } set { _Dialogue.InitialDirectory = value; } }
        public Boolean RestaurerLeDossierActif { get { return _Dialogue.RestoreDirectory; } set { _Dialogue.RestoreDirectory = value; } }
        public Boolean AfficherAide { get { return _Dialogue.ShowHelp; } set { _Dialogue.ShowHelp = value; } }
        public Boolean AfficherLectureSeule { get { return _Dialogue.ShowReadOnly; } set { _Dialogue.ShowReadOnly = value; } }
        public Boolean ExtensionsMultiple { get { return _Dialogue.SupportMultiDottedExtensions; } set { _Dialogue.SupportMultiDottedExtensions = value; } }

        public Boolean LectureSeuleSelectionne { get { return _Dialogue.ReadOnlyChecked; } set { _Dialogue.ReadOnlyChecked = value; } }

        #endregion

        #region "Méthodes"

        public String SelectionnerUnFichier(Boolean CheminComplet = true)
        {
            _Dialogue.Multiselect = false;

            if (_Dialogue.ShowDialog() == DialogResult.OK)
            {
                if (CheminComplet)
                    return _Dialogue.FileName;
                else
                    return _Dialogue.SafeFileName;
            }
            else
                return "";
        }

        public String[] SelectionnerPlusieursFichiers(Boolean CheminComplet = true)
        {
            _Dialogue.Multiselect = true;

            if (_Dialogue.ShowDialog() == DialogResult.OK)
            {
                if (CheminComplet)
                    return _Dialogue.FileNames;
                else
                    return _Dialogue.SafeFileNames;
            }
            else
                return null;
        }

        #endregion
    }
}
