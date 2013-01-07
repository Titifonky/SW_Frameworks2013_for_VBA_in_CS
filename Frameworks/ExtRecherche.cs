using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A8C91882-5820-11E2-A1E0-98046188709B")]
    public interface IExtRecherche
    {
        ExtComposant Composant { get;}
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
        Boolean RenvoyerComposantRacine { get; set; }
        Boolean Init(ExtComposant Composant);
        ArrayList Lancer(TypeFichier_e TypeComposant, String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.ExtRecherche")]
    public class ExtRecherche : IExtRecherche
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;

        private ExtComposant _Composant;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        private Boolean _RenvoyerComposantRacine = false;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtRecherche()
        {
        }

        #endregion

        #region "Propriétés"

        public ExtComposant Composant
        {
            get { return _Composant; }
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

        public Boolean RenvoyerComposantRacine
        {
            get { return _RenvoyerComposantRacine; }
            set { _RenvoyerComposantRacine = value; }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ExtComposant Composant)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!(Composant == null))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Composant = Composant;
                return true;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        private String NomCle(ExtComposant Composant)
        {
            String pNomCle = Composant.Modele.Chemin;
            if (_PrendreEnCompteConfig)
                pNomCle = pNomCle + " " + Composant.Configuration.Nom;

            return pNomCle;
        }

        private void RecListComposants(ExtComposant ComposantRacine, TypeFichier_e TypeComposant, Dictionary<String, ExtComposant> DicComposants, String NomComposant = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            foreach (ExtComposant Comp in ComposantRacine.ListComposantsEnfants(_PrendreEnCompteSupprime))
            {
                if (!Comp.EstExclu | _PrendreEnCompteExclus)
                {
                    if (Comp.Modele.TypeDuModele == TypeComposant)//&& Path.GetFileName(Comp.Modele.Chemin).Contains(NomComposant)
                    {
                        ExtComposant Composant = new ExtComposant();
                        
                        if (DicComposants.ContainsKey(NomCle(Comp)))
                        {
                            Composant = DicComposants[NomCle(Comp)];
                            Composant.Nb += 1;
                        }
                        else
                        {
                            Composant = Comp;
                            Composant.Nb = 1;
                            DicComposants.Add(NomCle(Composant), Composant);
                        }

                    }

                    if ((Comp.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (Comp.EstSupprime == false))
                    {
                        RecListComposants(Comp, TypeComposant, DicComposants, NomComposant);
                    }

                }
            }
        }

        internal List<ExtComposant> ListComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            Dictionary<String, ExtComposant> pDicComposants = new Dictionary<string, ExtComposant>();

            // On renvoi le composant de base seulement s'il a le meme type que ceux recherché (TypeComposant)
            if ((_RenvoyerComposantRacine == true) && (_Composant.Modele.TypeDuModele == TypeComposant))
                pDicComposants.Add(NomCle(_Composant), _Composant);

            // Si le composant est un assemblage contenant plusieurs composants, on renvoi la liste des composants recherchés
            if ((_Composant.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (_Composant.swComposant.IGetChildrenCount() > 0))
                RecListComposants(_Composant, TypeComposant, pDicComposants, NomComposant);

            // Nouvelle liste à renvoyer
            List<ExtComposant> pListeComposants = new List<ExtComposant>();

            // Si le dictionnaire n'est pas vide, on rempli la liste avec les valeurs du dictionnaire
            if (pDicComposants.Count > 0)
                pListeComposants = new List<ExtComposant>(pDicComposants.Values);

            // On trie et c'est parti
            pListeComposants.Sort();

            return pListeComposants;
        }

        public ArrayList Lancer(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtComposant> pListeComps = ListComposants(TypeComposant, NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion
    }
}
