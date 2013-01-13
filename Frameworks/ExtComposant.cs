using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("BF2ED17A-5820-11E2-8160-9F046188709B")]
    public interface IExtComposant
    {
        Component2 swComposant { get; }
        ExtModele Modele { get; }
        ExtConfiguration Configuration { get; }
        Boolean EstExclu { get; set; }
        Boolean EstSupprime { get; }
        int Nb { get; set; }
        ExtRecherche NouvelleRecherche { get; }
        ArrayList ComposantsEnfants(Boolean PrendreEnCompteSupprime = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C46318AE-5820-11E2-A863-A3046188709B")]
    [ProgId("Frameworks.ExtComposant")]
    public class ExtComposant : IExtComposant, IComparable<ExtComposant>, IComparer<ExtComposant>, IEquatable<ExtComposant>
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private Component2 _SwComposant;
        private ExtModele _Modele;
        private ExtConfiguration _Configuration;
        private int _Nb = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtComposant()
        {
        }

        #endregion

        #region "Propriétés"

        public Component2 swComposant { get { return _SwComposant; } }

        public ExtModele Modele { get { return _Modele; } }

        public ExtConfiguration Configuration { get { return _Configuration; } }

        public int Nb { get { return _Nb; } set { _Nb = value; } }

        public Boolean EstExclu
        {
            get
            {
                if (_SwComposant.ExcludeFromBOM != false)
                    return true;
                return false;
            }
            set { _SwComposant.ExcludeFromBOM = value; }
        }

        public Boolean EstSupprime
        {
            get
            {
                if (_SwComposant.IsSuppressed() != false)
                    return true;
                return false;
            }
        }

        /// <summary>
        /// Renvoi un nouvel objet Recherche
        /// </summary>
        public ExtRecherche NouvelleRecherche
        {
            get
            {
                ExtRecherche pNouvelleRecherche = new ExtRecherche();

                if (pNouvelleRecherche.Init(this))
                    return pNouvelleRecherche;

                return null;
            }
        }

        public Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Component2 SwComposant, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            // On teste si le Modele est valide
            if ((SwComposant != null) && (Modele != null) && Modele.EstInitialise)
            {
                // On valide l'initialisation avant de recupérer la configuration
                _EstInitialise = true;

                Configuration pSwConfig;

                // Si la configuration referencé est vide, on recupère la configuration active.
                if (String.IsNullOrEmpty(SwComposant.ReferencedConfiguration))
                    pSwConfig = Modele.SwModele.GetActiveConfiguration();
                else
                    pSwConfig = Modele.SwModele.GetConfigurationByName(SwComposant.ReferencedConfiguration);

                _Configuration = new ExtConfiguration();

                // Si la config est ok
                if (_Configuration.Init(pSwConfig, Modele))
                {
                    _SwComposant = SwComposant;
                    _Modele = Modele;
                    _Nb = 1;
                }
                else
                {
                    _EstInitialise = false;
                }
            }
            else // Sinon, on envoi pour le debug
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        internal List<ExtComposant> ListComposantsEnfants(Boolean PrendreEnCompteSupprime = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtComposant> Liste = new List<ExtComposant>();

            if (_SwComposant.IGetChildrenCount() == 0)
                return Liste;

            foreach (Component2 SwComposant in _SwComposant.GetChildren())
            {
                /// Si le composant est supprimé mais qu'on a decidé de le prendre en compte, c'est bon
                if ((SwComposant.IsSuppressed() == false) | PrendreEnCompteSupprime)
                {
                    // Pour intitialiser le composant correctement il faut un peu de bidouille
                    // sinon on à le droit à une belle reference circulaire
                    // Donc d'abord, on recherche le modele du SwComposant
                    ExtModele Modele = _Modele.SW.Modele(SwComposant.GetPathName());
                    // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
                    ExtComposant Composant = new ExtComposant();
                    Composant.Init(SwComposant, Modele);
                    // Et pour que les deux soit liés, on passe la ref du Composant que l'on vient de creer
                    // au modele. Comme ca, Modele.Composant pointe sur Composant et Composant.Modele pointe sur Modele,
                    // la boucle est bouclée
                    Modele.Composant = Composant;
                    Liste.Add(Composant);
                }
            }

            Liste.Sort();
            return Liste;

        }

        public ArrayList ComposantsEnfants(Boolean PrendreEnCompteSupprime = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtComposant> pListeComps = ListComposantsEnfants(PrendreEnCompteSupprime);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtComposant>.CompareTo(ExtComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom;
            String Nom2 = Comp.swComposant.GetPathName() + Comp.Configuration.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtComposant>.Compare(ExtComposant Comp1, ExtComposant Comp2)
        {
            String Nom1 = Comp1.swComposant.GetPathName() + Comp1.Configuration.Nom;
            String Nom2 = Comp2.swComposant.GetPathName() + Comp2.Configuration.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtComposant>.Equals(ExtComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom;
            String Nom2 = Comp.swComposant.GetPathName() + Comp.Configuration.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}