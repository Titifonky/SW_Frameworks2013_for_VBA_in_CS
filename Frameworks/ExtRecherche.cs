using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A8C91882-5820-11E2-A1E0-98046188709B")]
    public interface IExtRecherche
    {
        ExtModele Modele { get;}
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
        Boolean Init(ExtModele Modele);
        String NomCle(ExtComposant Composant);

    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.ExtRecherche")]
    public class ExtRecherche : IExtRecherche
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;

        private ExtModele _Modele;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        //private List<ExtComposant> _ListeComposants;

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

        public Boolean Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!(Modele == null))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Modele = Modele;
                return true;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        public String NomCle(ExtComposant Composant)
        {
            String pNomCle = Composant.Modele.Chemin;
            if (_PrendreEnCompteConfig)
                pNomCle = pNomCle + " " + Composant.Configuration.Nom;

            return pNomCle;
        }

        #endregion
    }
}
