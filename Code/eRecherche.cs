using SolidWorks.Interop.sldworks;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
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
        Boolean RenvoyerConfigComposantRacine { get; set; }
        ArrayList ListeComposants(TypeFichier_e TypeComposant, String NomComposant = "");
        ArrayList ListeFichiers(TypeFichier_e TypeComposant, String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.eRecherche")]
    public class eRecherche : IeRecherche
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eRecherche).Name;

        private Boolean _EstInitialise = false;

        private eComposant _Composant = null;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        private Boolean _SupprimerDoublons = true;
        private Boolean _RenvoyerComposantRacine = false;
        private Boolean _RenvoyerConfigComposantRacine = false;
        private static Double _IndexComposant = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public eRecherche() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtComposant.
        /// </summary>
        public eComposant Composant { get { Log.Methode(cNOMCLASSE); return _Composant; } }

        /// <summary>
        /// Filtre sur les configurations.
        /// </summary>
        public Boolean PrendreEnCompteConfig { get { Log.Methode(cNOMCLASSE); return _PrendreEnCompteConfig; } set { Log.Methode(cNOMCLASSE); _PrendreEnCompteConfig = value; } }

        /// <summary>
        /// Filtre sur les composants exclus.
        /// </summary>
        public Boolean PrendreEnCompteExclus { get { Log.Methode(cNOMCLASSE); return _PrendreEnCompteExclus; } set { Log.Methode(cNOMCLASSE); _PrendreEnCompteExclus = value; } }

        /// <summary>
        /// Filtre sur les composants supprimés.
        /// </summary>
        public Boolean PrendreEnCompteSupprime { get { Log.Methode(cNOMCLASSE); return _PrendreEnCompteSupprime; } set { Log.Methode(cNOMCLASSE); _PrendreEnCompteSupprime = value; } }

        /// <summary>
        /// Filtre sur les doublons.
        /// </summary>
        public Boolean SupprimerDoublons { get { Log.Methode(cNOMCLASSE); return _SupprimerDoublons; } set { Log.Methode(cNOMCLASSE); _SupprimerDoublons = value; } }

        /// <summary>
        /// Inclus le composant racine dans la liste,
        /// si celui ci est de même type que le type de composant filtré.
        /// </summary>
        public Boolean RenvoyerComposantRacine { get { Log.Methode(cNOMCLASSE); return _RenvoyerComposantRacine; } set { Log.Methode(cNOMCLASSE); _RenvoyerComposantRacine = value; } }

        public Boolean RenvoyerConfigComposantRacine { get { Log.Methode(cNOMCLASSE); return _RenvoyerConfigComposantRacine; } set { Log.Methode(cNOMCLASSE); _RenvoyerConfigComposantRacine = value; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet ExtRecherche
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((Composant != null) && Composant.EstInitialise)
            {
                _Composant = Composant;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
            Log.Methode(cNOMCLASSE);

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
        private void RecListerComposants(eComposant ComposantRacine, TypeFichier_e TypeComposant, Dictionary<String, eComposant> DicComposants, String NomComposant = "")
        {
            Log.Methode(cNOMCLASSE);

            foreach (eComposant pComp in ComposantRacine.ComposantsEnfants("", _PrendreEnCompteSupprime))
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

                        Log.Message("Clé : " + pCle);

                        // S'il est déjà dans le dico, on on rajoute 1
                        if (DicComposants.ContainsKey(pCle))
                        {
                            pComposant = DicComposants[pCle];
                            pComposant.Nb += 1;
                        }
                        // sinon on le rajoute
                        else
                        {
                            // Si on supprime les doublons et que l'on ne prend pas e compte les configs, on associe au modèle le composant racine
                            // et non le composant de l'assemblage
                            if (_SupprimerDoublons && !_PrendreEnCompteConfig)
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
                        RecListerComposants(pComp, TypeComposant, DicComposants, NomComposant);
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
        public ArrayList ListeComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            Log.Methode(cNOMCLASSE);

            Dictionary<String, eComposant> pDicComposants = new Dictionary<string, eComposant>();

            // On met l'index à 0
            _IndexComposant = 0;

            // On renvoi le composant de base
            if ((_RenvoyerComposantRacine == true) && TypeComposant.HasFlag(_Composant.Modele.TypeDuModele))
            {
                // On renvoi un composant par config
                if (_PrendreEnCompteConfig && _RenvoyerConfigComposantRacine)
                {
                    foreach (eConfiguration pConfig in _Composant.Modele.GestDeConfigurations.ListerLesConfigs(TypeConfig_e.cDeBase))
                    {
                        Component2 pSwComp = pConfig.SwConfiguration.GetRootComponent3(false);
                        eModele pModele = _Composant.Modele.SW.Modele(pSwComp.GetPathName());
                        eComposant pComp = new eComposant();
                        pComp.Init(pSwComp, pModele);
                        pModele.Composant = pComp;
                        pComp.Configuration = pConfig;
                        pDicComposants.Add(NomCle(pComp), pComp);
                    }
                }
                else
                    pDicComposants.Add(NomCle(_Composant), _Composant);
            }

            // Si le composant est un assemblage contenant plusieurs composants, on renvoi la liste des composants recherchés
            if ((_Composant.Modele.TypeDuModele == TypeFichier_e.cAssemblage) && (_Composant.SwComposant.IGetChildrenCount() > 0))
                RecListerComposants(_Composant, TypeComposant, pDicComposants, NomComposant);

            // Nouvelle liste à renvoyer
            ArrayList pListeComposants = new ArrayList();

            // Si le dictionnaire n'est pas vide, on rempli la liste avec les valeurs du dictionnaire
            if (pDicComposants.Count > 0)
                pListeComposants = new ArrayList(pDicComposants.Values);

            // On trie et c'est parti
            //pListeComposants.Sort();

            return pListeComposants;
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
            Log.Methode(cNOMCLASSE);

            ArrayList pListeComps = ListeComposants(TypeComposant, NomComposant);
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
