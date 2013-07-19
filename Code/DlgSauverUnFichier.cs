using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("8E8E5EA7-3594-4870-90B3-1C521304E4AE")]
    public interface IDlgSauverUnFichier
    {
        String Titre { get; set; }
        Boolean AjouterExtension { get; set; }
        Boolean TestFichierExiste { get; set; }
        Boolean TestCheminExiste { get; set; }
        Boolean DemanderAutorisationCreationFichier { get; set; }
        String ExtensionParDefaut { get; set; }
        String Filtre { get; set; }
        int IndexDuFiltre { get; set; }
        String DossierInitial { get; set; }
        Boolean RestaurerLeDossierActif { get; set; }
        Boolean AfficherAide { get; set; }
        Boolean ExtensionsMultiple { get; set; }

        String SauverUnFichier();
        String[] SauverPlusieursFichiers();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DC7FCAF0-3569-40CB-A281-5143BC990C8F")]
    [ProgId("Frameworks.DlgSauverUnFichier")]
    public class DlgSauverUnFichier : IDlgSauverUnFichier
    {
#region "Variables locales"

#endregion

#region "Constructeur\Destructeur"

        public DlgSauverUnFichier() { }

#endregion

#region "Propriétés"

        private SaveFileDialog _Dialogue = new SaveFileDialog();

        public String Titre { get { return _Dialogue.Title; } set { _Dialogue.Title = value; } }
        public Boolean AjouterExtension { get { return _Dialogue.AddExtension; } set { _Dialogue.AddExtension = value; } }
        public Boolean TestFichierExiste { get { return _Dialogue.CheckFileExists; } set { _Dialogue.CheckFileExists = value; } }
        public Boolean TestCheminExiste { get { return _Dialogue.CheckPathExists; } set { _Dialogue.CheckPathExists = value; } }
        public Boolean DemanderAutorisationCreationFichier { get { return _Dialogue.CreatePrompt; } set { _Dialogue.CreatePrompt = value; } }
        public String ExtensionParDefaut { get { return _Dialogue.DefaultExt; } set { _Dialogue.DefaultExt = value; } }
        public String Filtre { get { return _Dialogue.Filter; } set { _Dialogue.Filter = value; } }
        public int IndexDuFiltre { get { return _Dialogue.FilterIndex; } set { _Dialogue.FilterIndex = value; } }
        public String DossierInitial { get { return _Dialogue.InitialDirectory; } set { _Dialogue.InitialDirectory = value; } }
        public Boolean RestaurerLeDossierActif { get { return _Dialogue.RestoreDirectory; } set { _Dialogue.RestoreDirectory = value; } }
        public Boolean AfficherAide { get { return _Dialogue.ShowHelp; } set { _Dialogue.ShowHelp = value; } }
        public Boolean ExtensionsMultiple { get { return _Dialogue.SupportMultiDottedExtensions; } set { _Dialogue.SupportMultiDottedExtensions = value; } }


#endregion

#region "Méthodes"

        public String SauverUnFichier()
        {

            if (_Dialogue.ShowDialog() == DialogResult.OK)
                return _Dialogue.FileName;
            else
                return "";
        }

        public String[] SauverPlusieursFichiers()
        {
            if (_Dialogue.ShowDialog() == DialogResult.OK)
                return _Dialogue.FileNames;
            else
                return null;
        }

#endregion
    }
}
