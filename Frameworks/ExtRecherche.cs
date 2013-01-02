using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("928FDEF6-5529-11E2-A516-706D6188709B")]
    public interface IExtRecherche
    {
        ExtModele Modele { get; set; }
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9D886364-5529-11E2-8524-726D6188709B")]
    [ProgId("Frameworks.ExtRecherche")]
    public class ExtRecherche : IExtRecherche
    {
        #region "Variables locales"

        private ExtModele _Modele;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtRecherche()
        {
        }

        #endregion

        #region "Propriétés"

        public ExtModele Modele
        {
            get { return _Modele; }
            set { _Modele = value; }
        }

        public Boolean PrendreEnCompteConfig
        {
            get { return _PrendreEnCompteConfig; }
            set { _PrendreEnCompteConfig = value; }
        }

        public Boolean PrendreEnCompteExclus
        {
            get { return _PrendreEnCompteExclus; }
            set { _PrendreEnCompteExclus = value; }
        }

        public Boolean PrendreEnCompteSupprime
        {
            get { return _PrendreEnCompteSupprime; }
            set { _PrendreEnCompteSupprime = value; }
        }

        #endregion

        #region "Méthodes"
        #endregion
    }
}
