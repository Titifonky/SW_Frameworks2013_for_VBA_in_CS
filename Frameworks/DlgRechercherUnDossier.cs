using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("64A83C8C-997C-4266-83EC-E6B99A58FE94")]
    public interface IDlgRechercherUnDossier
    {
        String Description { get; set; }
        Environment.SpecialFolder DossierRacine { get; set; }
        Boolean BoutonNouveauDossier { get; set; }
        String SelectionnerUnDossier();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BADDD2D2-9D97-4D6B-8F49-C9BC161E3E06")]
    [ProgId("Frameworks.DlgRechercherUnDossier")]
    public class DlgRechercherUnDossier : IDlgRechercherUnDossier
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public DlgRechercherUnDossier() { }

        #endregion

        #region "Propriétés"

        private FolderBrowserDialog _Dialogue = new FolderBrowserDialog();

        public String Description { get { return _Dialogue.Description; } set { _Dialogue.Description = value; } }
        public Environment.SpecialFolder DossierRacine { get { return _Dialogue.RootFolder; } set { _Dialogue.RootFolder = value; } }
        public Boolean BoutonNouveauDossier { get { return _Dialogue.ShowNewFolderButton; } set { _Dialogue.ShowNewFolderButton = value; } }

        #endregion

        #region "Méthodes"

        public String SelectionnerUnDossier()
        {
            if (_Dialogue.ShowDialog() == DialogResult.OK)
                return _Dialogue.SelectedPath;
            
            return "";
        }

        #endregion
    }
}
