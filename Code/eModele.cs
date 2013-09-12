﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("66AE684E-5820-11E2-BCFB-5D046188709B")]
    public interface IeModele
    {
        ModelDoc2 SwModele { get; }
        eSldWorks SW { get; }
        eComposant Composant { get; }
        eAssemblage Assemblage { get; }
        ePiece Piece { get; }
        eDessin Dessin { get; }
        eGestDeConfigurations GestDeConfigurations { get; }
        eGestDeProprietes GestDeProprietes { get; }
        eGestDeSelection GestDeSelection { get; }
        TypeFichier_e TypeDuModele { get; }
        eFichierSW FichierSw { get; }
        Boolean EstActif { get; set; }
        Boolean ActiverInterfaceUtilisateur { get; set; }
        void Activer(swRebuildOnActivation_e Reconstruire = swRebuildOnActivation_e.swUserDecision);
        void Sauver();
        void Fermer();
        void Redessiner();
        void Reconstruire();
        void ForcerAToutReconstruire();
        void ZoomEtendu();
        void EffacerLesSelections();
        ArrayList ListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false);
        eFonction DerniereFonction();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6AFCE66E-5820-11E2-B651-77046188709B")]
    [ProgId("Frameworks.eModele")]
    public class eModele : IeModele, IComparable<eModele>, IComparer<eModele>, IEquatable<eModele>
    {
#region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ModelDoc2 _SwModele = null;
        private eSldWorks _SW = null;
        private eComposant _Composant = null;
        private eAssemblage _Assemblage = null;
        private ePiece _Piece = null;
        private eDessin _Dessin = null;
        private eGestDeSelection _GestDeSelection = null;
        private eFichierSW _FichierSw = null;
        private int Erreur = 0;
        private int Warning = 0;

#endregion

#region "Constructeur\Destructeur"

        public eModele() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le modele ModleDoc2 associé.
        /// </summary>
        public ModelDoc2 SwModele
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _SwModele;
            }
        }

        /// <summary>
        /// Retourne le parent ExtSldWorks.
        /// </summary>
        public eSldWorks SW { get { Debug.Print(MethodBase.GetCurrentMethod());  return _SW; } }

        /// <summary>
        /// Retourne le composant ExtComposant lié au modele.
        /// </summary>
        public eComposant Composant
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod()); 
                return _Composant;
            }
            internal set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                _Composant = value;
            }
        }

        /// <summary>
        /// Retourne l'assemblage ExtAssemblage si celui ci est valide.
        /// </summary>
        public eAssemblage Assemblage
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Assemblage == null)
                {
                    _Assemblage = new eAssemblage();
                    _Assemblage.Init(this);
                }

                if (_Assemblage.EstInitialise)
                    return _Assemblage;

                return null;
            }
        }

        /// <summary>
        /// Retourne la pièce ExtPiece si celui ci est valide.
        /// </summary>
        public ePiece Piece
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Piece == null)
                {
                    _Piece = new ePiece();
                    _Piece.Init(this);
                }

                if (_Piece.EstInitialise)
                    return _Piece;

                return null;
            }
        }

        /// <summary>
        /// Retourne le dessin ExtDessin si celui ci est valide.
        /// </summary>
        public eDessin Dessin
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Dessin == null)
                {
                    _Dessin = new eDessin();
                    _Dessin.Init(this);
                }

                if (_Dessin.EstInitialise)
                    return _Dessin;
                
                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de configuration GestDeConfigurations.
        /// </summary>
        public eGestDeConfigurations GestDeConfigurations
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eGestDeConfigurations pGestDeConfigurations = new eGestDeConfigurations();

                if (pGestDeConfigurations.Init(this))
                    return pGestDeConfigurations;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public eGestDeProprietes GestDeProprietes
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eGestDeProprietes pGestDeProprietes = new eGestDeProprietes();

                if (pGestDeProprietes.Init(SwModele.Extension.get_CustomPropertyManager(""), this))
                    return pGestDeProprietes;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public eGestDeSelection GestDeSelection
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_GestDeSelection == null)
                {
                    _GestDeSelection = new eGestDeSelection();
                    _GestDeSelection.Init(SwModele.SelectionManager, this);
                }

                if (_GestDeSelection.EstInitialise)
                    return _GestDeSelection;

                return null;
            }
        }

        /// <summary>
        /// Retourne le type du modele.
        /// </summary>
        public TypeFichier_e TypeDuModele
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                switch (_SwModele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        return TypeFichier_e.cAssemblage;

                    case (int)swDocumentTypes_e.swDocPART:
                        return TypeFichier_e.cPiece;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        return TypeFichier_e.cDessin;

                    default:
                        return 0;
                }
            }
        }

        /// <summary>
        /// Retourne l'objet ExtFichierSw
        /// </summary>
        public eFichierSW FichierSw
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (Composant != null)
                {
                    _FichierSw.Nb = Composant.Nb;
                    if (Composant.SwComposant.IsRoot())
                        _FichierSw.Configuration = GestDeConfigurations.ConfigurationActive.Nom;
                    else
                        _FichierSw.Configuration = Composant.Configuration.Nom;

                }
                return _FichierSw;
            }
        }

        public Boolean EstActif
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (SW.Modele().Equals(this))
                    return true;

                return false;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (value)
                {
                    Activer();
                }
            }
        }

        /// <summary>
        /// Activer ou désactiver les éléments de l'interface pour accélérer les macros
        /// </summary>
        public Boolean ActiverInterfaceUtilisateur
        {
            get
            {
                return _SwModele.FeatureManager.EnableFeatureTree && _SwModele.FeatureManager.EnableFeatureTreeWindow && _SwModele.ConfigurationManager.EnableConfigurationTree;
            }
            set
            {
                _SwModele.FeatureManager.EnableFeatureTree = value;
                _SwModele.FeatureManager.EnableFeatureTreeWindow = value;
                _SwModele.ConfigurationManager.EnableConfigurationTree = value;
                //_SwModele.Lock();

                if (value)
                {
                    _SwModele.FeatureManager.UpdateFeatureTree();
                    //_SwModele.UnLock();
                }
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtModele.
        /// </summary>
        /// <param name="SwModele"></param>
        /// <param name="Sw"></param>
        /// <returns></returns>
        internal Boolean Init(ModelDoc2 SwModele, eSldWorks Sw)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwModele != null) && (Sw != null) && Sw.EstInitialise)
            {
                _SwModele = SwModele;
                _SW = Sw;
                Debug.Print(_SwModele.GetPathName());

                // On valide l'initialisation
                _EstInitialise = true;

                // On créer l'objet ExtFichierSw associé
                _FichierSw = new eFichierSW();
                if (_FichierSw.Init(Sw))
                    _FichierSw.Chemin = _SwModele.GetPathName();
                else
                    _EstInitialise = false;

                // Si c'est un assemblage ou une pièce, on va chercher le composant associé
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    Debug.Print("Referencement de la configuration");
                    Configuration pConfigActive = _SwModele.ConfigurationManager.ActiveConfiguration;
                    _FichierSw.Configuration = pConfigActive.Name;

                    Debug.Print("Referencement du composant");
                    _Composant = new eComposant();
                    if (_Composant.Init(pConfigActive.GetRootComponent3(false), this) == false)
                        _EstInitialise = false;
                }
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        internal void ReinitialiserComposant()
        {
            // Si c'est un assemblage ou une pièce, on va chercher le composant associé
            if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
            {
                Debug.Print("Referencement de la configuration");
                Configuration pConfigActive = _SwModele.ConfigurationManager.ActiveConfiguration;
                _FichierSw.Configuration = pConfigActive.Name;

                Debug.Print("Referencement du composant");
                _Composant = new eComposant();
                if (_Composant.Init(pConfigActive.GetRootComponent3(false), this) == false)
                    _EstInitialise = false;
                else
                    _Composant.Configuration = GestDeConfigurations.ConfigurationActive;
            }
        }

        /// <summary>
        /// Active le modele et le met au premier plan.
        /// </summary>
        public void Activer(swRebuildOnActivation_e Reconstruire = swRebuildOnActivation_e.swUserDecision)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

#if SW2012
            _SW.SwSW.ActivateDoc3(SwModele.GetPathName(), true, (int)Reconstruire, Erreur);
#else
            _SW.SwSW.ActivateDoc2(SwModele.GetPathName(), true, Erreur);
#endif
            ZoomEtendu();
        }

        /// <summary>
        /// Sauve le modele.
        /// </summary>
        public void Sauver()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            SwModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
            _FichierSw.Chemin = SwModele.GetPathName();
        }

        /// <summary>
        /// Ferme le modele.
        /// </summary>
        public void Fermer()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            _SW.SwSW.CloseDoc(SwModele.GetPathName());
        }

        /// <summary>
        /// Redessine le modele.
        /// </summary>
        public void Redessiner()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            SwModele.ActiveView.GraphicsRedraw();
        }

        /// <summary>
        /// Reconstruit le modele.
        /// </summary>
        public void Reconstruire()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            SwModele.EditRebuild3();
        }

        /// <summary>
        /// Force à tout reconstruire.
        /// </summary>
        public void ForcerAToutReconstruire()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            SwModele.ForceRebuild3(false);
        }

        /// <summary>
        /// Zoom étendu du modele.
        /// </summary>
        public void ZoomEtendu()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            SwModele.ViewZoomtofit2();
        }

        /// <summary>
        /// Effacer les selections
        /// </summary>
        public void EffacerLesSelections()
        {
            _SwModele.ClearSelection2(true);
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des fonctions filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<eFonction> ListListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = new List<eFonction>();

            Feature pSwFonction = _SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                eFonction pFonction = new eFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher))
                    && (Regex.IsMatch(pSwFonction.GetTypeName2(), TypeDeLaFonction))
                    && pFonction.Init(pSwFonction, this)
                    && !(pListeFonctions.Contains(pFonction))
                    )
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        eFonction pSousFonction = new eFonction();

                        if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher))
                            && (Regex.IsMatch(pSwSousFonction.GetTypeName2(), TypeDeLaFonction))
                            && pSousFonction.Init(pSwSousFonction, this)
                            && !(pListeFonctions.Contains(pSousFonction))
                            )
                            pListeFonctions.Add(pSousFonction);

                        pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                    }
                }

                pSwFonction = pSwFonction.GetNextFeature();
            }

            return pListeFonctions;
        }

        /// <summary>
        /// Renvoi la liste des fonctions filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        public ArrayList ListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        /// <summary>
        /// Renvoi la dernière fonction crée
        /// </summary>
        /// <returns></returns>
        public eFonction DerniereFonction()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            swFeatureType_e TypeF = new swFeatureType_e();
            int i = 0;
            Feature pSwFonc = _SwModele.FeatureByPositionReverse(i);


            while (pSwFonc != null)
            {
                eFonction pFonc = new eFonction();
                if ((pSwFonc.GetTypeName2() != TypeF.swTnFlatPattern) && pFonc.Init(pSwFonc, this))
                    return pFonc;

                i++;
                pSwFonc = _SwModele.FeatureByPositionReverse(i);
            }

            return null;
        }

        internal TreeControlItem GestDeFonction_NoeudRacine()
        {
#if SW2012
            return SwModele.FeatureManager.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneTop);
#else
            return SwModele.FeatureManager.GetFeatureTreeRootItem();
#endif
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(eModele Modele)
        {
            return _SwModele.GetPathName().CompareTo(Modele.SwModele.GetPathName());
        }

        public int Compare(eModele Modele1, eModele Modele2)
        {
            return Modele1.SwModele.GetPathName().CompareTo(Modele2.SwModele.GetPathName());
        }

        public Boolean Equals(eModele Modele)
        {
            return _SwModele.GetPathName().Equals(Modele.SwModele.GetPathName());
        }

#endregion
    }
}
