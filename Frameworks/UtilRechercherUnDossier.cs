using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("64A83C8C-997C-4266-83EC-E6B99A58FE94")]
    public interface IUtilRechercherUnDossier
    {
        String Titre { get;  set; }
        Environment.SpecialFolder DossierRacine { get; set; }
        String DossierPreSelectionne { get; set; }
        Boolean BoutonNouveauDossier { get; set; }
        String SelectionnerFichier();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BADDD2D2-9D97-4D6B-8F49-C9BC161E3E06")]
    [ProgId("Frameworks.UtilRechercherUnDossier")]
    class UtilRechercherUnDossier : IUtilRechercherUnDossier
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public UtilRechercherUnDossier() { }

        #endregion

        #region "Propriétés"

        public String Titre { get; set; }
        public Environment.SpecialFolder DossierRacine { get; set; }
        public String DossierPreSelectionne { get; set; }
        public Boolean BoutonNouveauDossier { get; set; }

        #endregion

        #region "Méthodes"

        public String SelectionnerFichier()
        {
            return "OK";
        }

        #endregion
    }
}
