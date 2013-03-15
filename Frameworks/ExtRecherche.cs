using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;

/////////////////////////// Implementation terminée ///////////////////////////

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A8C91882-5820-11E2-A1E0-98046188709B")]
    public interface IExtRecherche
    {
        ExtComposant Composant { get; }
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
        Boolean SupprimerDoublons { get; set; }
        Boolean RenvoyerComposantRacine { get; set; }
        ArrayList Lancer(TypeFichier_e TypeComposant, String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.ExtRecherche")]
    public class ExtRecherche : IExtRecherche
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ExtComposant _Composant;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        private Boolean _SupprimerDoublons = true;
        private Boolean _RenvoyerComposantRacine = false;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtRecherche() { }

        #endregion

        #region "Propriétés"

        public ExtComposant Composant { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Composant; } }

        public Boolean PrendreEnCompteConfig { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteConfig; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteConfig = value; } }

        public Boolean PrendreEnCompteExclus { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteExclus; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteExclus = value; } }

        public Boolean PrendreEnCompteSupprime { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteSupprime; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteSupprime = value; } }

        public Boolean SupprimerDoublons { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SupprimerDoublons; } set { Debug.Info(MethodBase.GetCurrentMethod()); _SupprimerDoublons = value; } }

        public Boolean RenvoyerComposantRacine { get { Debug.Info(MethodBase.GetCurrentMethod());  return _RenvoyerComposantRacine; } set { Debug.Info(MethodBase.GetCurrentMethod());  _RenvoyerComposantRacine = value; } }

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtComposant Composant)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Composant != null) && Composant.EstInitialise)
            {
                _Composant = Composant;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        private String NomCle(ExtComposant Composant)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            
            String pNomCle = "";
            if ((Composant != null) && Composant.EstInitialise)
            {
                pNomCle = Composant.Modele.Chemin;
                // Si on prend en compte les configs, on rajoute le nom de la config dans la clé
                if (_PrendreEnCompteConfig)
                    pNomCle += "_" + Composant.Configuration.Nom;

                // Si on garde toutes les occurences, on rajoute un nombre aléatoire à la fin de la clé
                // Comme ça, ils sont tous différents
                if (!_SupprimerDoublons)
                    pNomCle += "_" + (new Random()).Next();
            }

            return pNomCle;
        }

        private void RecListListerComposants(ExtComposant ComposantRacine, TypeFichier_e TypeComposant, Dictionary<String, ExtComposant> DicComposants, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            foreach (ExtComposant pComp in ComposantRacine.ListComposantsEnfants("",_PrendreEnCompteSupprime))
            {
                // Operateur "Implique" sur la propriété EstExclu
                if (!pComp.EstExclu | _PrendreEnCompteExclus)
                {
                    // Attention, Regex.IsMatch peut renvoyer une erreur si NomComposant est égal à un caractère spécial
                    // du style "*" ou "[" et autre. Pb à corriger.
                    if ( Convert.ToBoolean(pComp.Modele.TypeDuModele & TypeComposant) && Regex.IsMatch(pComp.Modele.Chemin, NomComposant))
                    {
                        ExtComposant pComposant = new ExtComposant();

                        // S'il est déjà dans le dico, on on rajoute 1
                        if (_SupprimerDoublons && DicComposants.ContainsKey(NomCle(pComp)))
                        {
                            pComposant = DicComposants[NomCle(pComp)];
                            pComposant.Nb += 1;
                        }
                        // sinon on le rajoute
                        else
                        {
                            pComposant = pComp;
                            pComposant.Nb = 1;
                            DicComposants.Add(NomCle(pComposant), pComposant);
                        }

                    }

                    // Si c'est un assemblage et qu'il n'est pas supprimé, on scan
                    if ((pComp.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (pComp.EstSupprime == false))
                    {
                        RecListListerComposants(pComp, TypeComposant, DicComposants, NomComposant);
                    }

                }
            }
        }

        internal List<ExtComposant> ListListerComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Dictionary<String, ExtComposant> pDicComposants = new Dictionary<string, ExtComposant>();

            // On renvoi le composant de base seulement s'il a le meme type que ceux recherché (TypeComposant)
            if ((_RenvoyerComposantRacine == true) && Convert.ToBoolean(_Composant.Modele.TypeDuModele & TypeComposant))
                pDicComposants.Add(NomCle(_Composant), _Composant);

            // Si le composant est un assemblage contenant plusieurs composants, on renvoi la liste des composants recherchés
            if ((_Composant.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (_Composant.SwComposant.IGetChildrenCount() > 0))
                RecListListerComposants(_Composant, TypeComposant, pDicComposants, NomComposant);

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
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtComposant> pListeComps = ListListerComposants(TypeComposant, NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion
    }
}
