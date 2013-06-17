using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A8C91882-5820-11E2-A1E0-98046188709B")]
    public interface IeRecherche
    {
        eComposant Composant { get; }
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
        Boolean SupprimerDoublons { get; set; }
        Boolean RenvoyerComposantRacine { get; set; }
        ArrayList ListeComposants(TypeFichier_e TypeComposant, String NomComposant = "");
        ArrayList ListeFichiers(TypeFichier_e TypeComposant, String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.eRecherche")]
    public class eRecherche : IeRecherche
    {
#region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eComposant _Composant = null;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        private Boolean _SupprimerDoublons = true;
        private Boolean _RenvoyerComposantRacine = false;
        private static Double _IndexComposant = 0;

#endregion

#region "Constructeur\Destructeur"

        public eRecherche() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtComposant.
        /// </summary>
        public eComposant Composant { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Composant; } }

        /// <summary>
        /// Filtre sur les configurations.
        /// </summary>
        public Boolean PrendreEnCompteConfig { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteConfig; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteConfig = value; } }

        /// <summary>
        /// Filtre sur les composants exclus.
        /// </summary>
        public Boolean PrendreEnCompteExclus { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteExclus; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteExclus = value; } }

        /// <summary>
        /// Filtre sur les composants supprimés.
        /// </summary>
        public Boolean PrendreEnCompteSupprime { get { Debug.Info(MethodBase.GetCurrentMethod());  return _PrendreEnCompteSupprime; } set { Debug.Info(MethodBase.GetCurrentMethod());  _PrendreEnCompteSupprime = value; } }

        /// <summary>
        /// Filtre sur les doublons.
        /// </summary>
        public Boolean SupprimerDoublons { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SupprimerDoublons; } set { Debug.Info(MethodBase.GetCurrentMethod()); _SupprimerDoublons = value; } }

        /// <summary>
        /// Inclus le composant racine dans la liste,
        /// si celui ci est de même type que le type de composant filtré.
        /// </summary>
        public Boolean RenvoyerComposantRacine { get { Debug.Info(MethodBase.GetCurrentMethod());  return _RenvoyerComposantRacine; } set { Debug.Info(MethodBase.GetCurrentMethod());  _RenvoyerComposantRacine = value; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet ExtRecherche
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtRecherche.
        /// </summary>
        /// <param name="Composant"></param>
        /// <returns></returns>
        internal Boolean Init(eComposant Composant)
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

        /// <summary>
        /// Méthode privée.
        /// Retourne la clé de tri suivant les filtres activé.
        /// </summary>
        /// <param name="Composant"></param>
        /// <returns></returns>
        private String NomCle(eComposant Composant)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            
            String pNomCle = "";
            if ((Composant != null) && Composant.EstInitialise)
            {
                pNomCle = Composant.Modele.FichierSw.Chemin;
                // Si on prend en compte les configs, on rajoute le nom de la config dans la clé
                if (_PrendreEnCompteConfig)
                    pNomCle += "_" + Composant.Configuration.Nom;

                // Si on garde toutes les occurences, on rajoute un index pour eviter les doublons de clés
                if (!_SupprimerDoublons)
                    pNomCle += "_" + _IndexComposant++;
            }

            return pNomCle;
        }

        /// <summary>
        /// Méthode privée.
        /// Méthode récursive permettant de traverser tous les composants d'une pièce ou d'un assemblage
        /// </summary>
        /// <param name="ComposantRacine"></param>
        /// <param name="TypeComposant"></param>
        /// <param name="DicComposants"></param>
        /// <param name="NomComposant"></param>
        private void RecListListerComposants(eComposant ComposantRacine, TypeFichier_e TypeComposant, Dictionary<String, eComposant> DicComposants, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            foreach (eComposant pComp in ComposantRacine.ListComposantsEnfants("",_PrendreEnCompteSupprime))
            {
                // Operateur "Implique" sur la propriété EstExclu
                if (!pComp.EstExclu | _PrendreEnCompteExclus)
                {
                    // Attention, Regex.IsMatch peut renvoyer une erreur si NomComposant est égal à un caractère spécial
                    // du style "*" ou "[" et autre. Pb à corriger.
                    if (TypeComposant.HasFlag(pComp.Modele.TypeDuModele) && Regex.IsMatch(pComp.Modele.FichierSw.Chemin, NomComposant))
                    {
                        eComposant pComposant = new eComposant();
                        String pCle = NomCle(pComp);

                        Debug.Info("Clé : " + pCle);

                        // S'il est déjà dans le dico, on on rajoute 1
                        if (DicComposants.ContainsKey(pCle))
                        {
                            pComposant = DicComposants[pCle];
                            pComposant.Nb += 1;
                        }
                        // sinon on le rajoute
                        else
                        {
                            // Si on supprime les doublons, on doit associer au modèle le composant racine
                            // et non le composant de l'assemblage
                            if (_SupprimerDoublons)
                            {
                                eModele pModele = pComp.Modele;
                                pModele.ReinitialiserComposant();
                                pComposant = pModele.Composant;
                            }
                            else
                            {
                                pComposant = pComp;
                            }

                            pComposant.Nb = 1;
                            DicComposants.Add(pCle, pComposant);
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

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des composants filtrées par les arguments
        /// </summary>
        /// <param name="TypeComposant"></param>
        /// <param name="NomComposant"></param>
        /// <returns></returns>
        internal List<eComposant> ListListerComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Dictionary<String, eComposant> pDicComposants = new Dictionary<string, eComposant>();

            // On met l'index à 0
            _IndexComposant = 0;

            // On renvoi le composant de base
            if ((_RenvoyerComposantRacine == true) && TypeComposant.HasFlag(_Composant.Modele.TypeDuModele))
                pDicComposants.Add(NomCle(_Composant), _Composant);

            // Si le composant est un assemblage contenant plusieurs composants, on renvoi la liste des composants recherchés
            if ((_Composant.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (_Composant.SwComposant.IGetChildrenCount() > 0))
                RecListListerComposants(_Composant, TypeComposant, pDicComposants, NomComposant);

            // Nouvelle liste à renvoyer
            List<eComposant> pListeComposants = new List<eComposant>();

            // Si le dictionnaire n'est pas vide, on rempli la liste avec les valeurs du dictionnaire
            if (pDicComposants.Count > 0)
                pListeComposants = new List<eComposant>(pDicComposants.Values);

            // On trie et c'est parti
            pListeComposants.Sort();

            return pListeComposants;
        }

        /// <summary>
        /// Renvoi la liste des composants filtrées par les arguments
        /// </summary>
        /// <param name="TypeComposant"></param>
        /// <param name="NomComposant"></param>
        /// <returns></returns>
        public ArrayList ListeComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eComposant> pListeComps = ListListerComposants(TypeComposant, NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        /// <summary>
        /// Renvoi la liste des ExtFichierSW filtrées par les arguments
        /// Permet de lister les fichiers sans faire de liens avec les composants SW et donc de pouvoir
        /// fermer le composant racine pour ensuite travailler sur les fichiers.
        /// Cela permet d'accélérer le traitement des fichiers.
        /// </summary>
        /// <param name="TypeComposant"></param>
        /// <param name="NomComposant"></param>
        /// <returns></returns>
        public ArrayList ListeFichiers(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eComposant> pListeComps = ListListerComposants(TypeComposant, NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
            {
                foreach (eComposant pComp in pListeComps)
                {
                    eFichierSW pFichier = new eFichierSW();
                    if (pFichier.Init(_Composant.Modele.SW))
                    {
                        pFichier.Chemin = pComp.Modele.FichierSw.Chemin;
                        pFichier.Configuration = pComp.Configuration.Nom;
                        pFichier.Nb = pComp.Nb;
                        pArrayComps.Add(pFichier);
                    }
                }
            }

            return pArrayComps;
        }

#endregion
    }
}
