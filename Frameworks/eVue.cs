using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Reflection;
using System.Collections.Generic;
using SolidWorks.Interop.swconst;
using System.Text.RegularExpressions;
using System.Collections;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("7EECBE39-3E38-49D7-A617-2E3AFEF915ED")]
    public interface IeVue
    {
        View SwVue { get; }
        eFeuille Feuille { get; }
        String Nom { get; set; }
        eModele ModeleDeReference { get; }
        eConfiguration ConfigurationDeReference { get; }
        eDimensionVue Dimensions { get; }
        Boolean AfficherLignesDePliage { set; }
        Boolean AfficherNotesDePliage { get; set; }
        void Selectionner(Boolean Ajouter = false);
        void Supprimer();
        //ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0A46F613-F061-4FC2-8DAD-F4EA5BBBBD8E")]
    [ProgId("Frameworks.eVue")]
    public class eVue : IeVue
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eFeuille _Feuille = null;
        private View _SwVue = null;
        #endregion

        #region "Constructeur\Destructeur"

        public eVue() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet View associé.
        /// </summary>
        public View SwVue { get { Debug.Print(MethodBase.GetCurrentMethod()); return _SwVue; } }

        /// <summary>
        /// Retourne le parent ExtFeuille.
        /// </summary>
        public eFeuille Feuille { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Feuille; } }

        /// <summary>
        /// Retourne ou défini le nom de la vue.
        /// </summary>
        public String Nom { get { Debug.Print(MethodBase.GetCurrentMethod()); return _SwVue.GetName2(); } set { Debug.Print(MethodBase.GetCurrentMethod()); _SwVue.SetName2(value); } }

        /// <summary>
        /// Retourne le modele ExtModele référencé par la vue.
        /// </summary>
        public eModele ModeleDeReference
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eModele pModele = new eModele();
                if (pModele.Init(_SwVue.ReferencedDocument, _Feuille.Dessin.Modele.SW))
                {
                    DrawingComponent pSwDwgComp = SwVue.RootDrawingComponent;
                    Component2 pSwComp = pSwDwgComp.Component;
                    if (pSwDwgComp.IsRoot())
                    {
                        Configuration pSwConfig = _SwVue.ReferencedDocument.GetConfigurationByName(_SwVue.ReferencedConfiguration);
                        pSwComp = pSwConfig.GetRootComponent3(false);
                    }
                    eComposant pComp = new eComposant();
                    pComp.Init(pSwComp, pModele);
                    pModele.Composant = pComp;
                    return pModele;
                }

                return null;
            }
        }

        /// <summary>
        /// Retourne la configuration ExtConfiguration référencée par la vue.
        /// </summary>
        public eConfiguration ConfigurationDeReference
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eConfiguration pConfig = new eConfiguration();
                if (pConfig.Init(_SwVue.ReferencedDocument.GetConfigurationByName(_SwVue.ReferencedConfiguration), ModeleDeReference))
                    return pConfig;

                return null;
            }
        }

        /// <summary>
        /// Retourne les dimensions de la vue.
        /// </summary>
        public eDimensionVue Dimensions
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eDimensionVue pDimensions = new eDimensionVue();

                if (pDimensions.Init(this))
                    return pDimensions;

                return null;
            }
        }

        /// <summary>
        /// Afficher ou masquer les lignes de pliage
        /// </summary>
        public Boolean AfficherLignesDePliage
        {
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                ModelDoc2 pSwModele = Feuille.Dessin.Modele.SwModele;
                eConfiguration pConfig = ConfigurationDeReference;

                if (pConfig.Est(TypeConfig_e.cDepliee))
                {
                    foreach (eFonction pFonction in ListListeDesFonctionsDeArbre("", "^FlatPattern$", true))
                    {
                        eFonction pFonctionEsquisse = pFonction.ListListeDesSousFonctions()[0];
                        if (pFonctionEsquisse != null)
                        {
                            String T;
                            String pNomPourSelection = pFonctionEsquisse.SwFonction.GetNameForSelection(out T);
                            pSwModele.Extension.SelectByID2(pNomPourSelection, T, 0, 0, 0, false, 0, null, 0);

                            // On affiche les lignes seulement si c'est la bonne fonction
                            // Sinon, on cache tout
                            if (value && (pFonction.Etat(pConfig) == EtatFonction_e.cActivee))
                                pSwModele.UnblankSketch();
                            else
                                pSwModele.BlankSketch();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Afficher ou masquer les notes de pliage
        /// </summary>
        public Boolean AfficherNotesDePliage
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _SwVue.ShowSheetMetalBendNotes;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                _SwVue.ShowSheetMetalBendNotes = value;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtVue.
        /// </summary>
        /// <param name="SwVue"></param>
        /// <param name="Feuille"></param>
        /// <returns></returns>
        internal Boolean Init(View SwVue, eFeuille Feuille)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwVue != null) && (Feuille != null) && Feuille.EstInitialise)
            {
                _Feuille = Feuille;
                _SwVue = SwVue;

                Debug.Print(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtVue.
        /// </summary>
        /// <param name="SwVue"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(View SwVue, eModele Modele)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwVue != null) && (Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                eFeuille Feuille = new eFeuille();
                if (Feuille.Init(SwVue.Sheet, Modele.Dessin))
                {
                    _Feuille = Feuille;
                    _SwVue = SwVue;

                    Debug.Print(this.Nom);
                    _EstInitialise = true;
                }
                else
                {
                    Debug.Print("!!!!! Erreur d'initialisation");
                }
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Supprimer la vue
        /// </summary>
        public void Supprimer()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            Selectionner();
            eModele pModeleDessin = Feuille.Dessin.Modele;
            pModeleDessin.Activer();
            pModeleDessin.SwModele.DeleteSelection(false);
            pModeleDessin.SwModele.ClearSelection2(true);
        }

        /// <summary>
        /// Selectionner la vue
        /// </summary>
        /// <param name="Ajouter"></param>
        public void Selectionner(Boolean Ajouter = false)
        {
            eModele pModeleDessin = Feuille.Dessin.Modele;
            pModeleDessin.Activer();

            ePoint pCentre = Dimensions.Centre;
            pModeleDessin.SwModele.Extension.SelectByID2(Nom, "DRAWINGVIEW", pCentre.X, pCentre.Y, pCentre.Z, Ajouter, -1, null, 0);
        }

        /// <summary>
        /// Scanner les fonctions de la vue à partir du FeatureManager
        /// </summary>
        /// <param name="Noeud"></param>
        /// <param name="ListeFonctions"></param>
        /// <param name="NomARechercher"></param>
        /// <param name="TypeDeLaFonction"></param>
        /// <param name="AvecLesSousFonctions"></param>
        private void ScannerFonctionsFeatureManager(TreeControlItem Noeud, List<eFonction> ListeFonctions, String NomARechercher, String TypeDeLaFonction, Boolean AvecLesSousFonctions)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            TreeControlItem pNoeud = Noeud.GetFirstChild();

            while (pNoeud != null)
            {
                eFonction pFonction = new eFonction();
                if (pNoeud.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Feature)
                {
                    if (pFonction.Init(pNoeud.Object, ModeleDeReference)
                        && Regex.IsMatch(pFonction.Nom, NomARechercher)
                        && Regex.IsMatch(pFonction.TypeDeLaFonction, TypeDeLaFonction))
                    {
                        ListeFonctions.Add(pFonction);
                    }
                }
                // On scanne dans tous les cas le dossier Tôlerie et le dossier Etat déplié 
                if (AvecLesSousFonctions ||
                    (pFonction.EstInitialise && (pFonction.TypeDeLaFonction == "TemplateSheetMetal"))
                    || (pFonction.EstInitialise && (pFonction.TypeDeLaFonction == "TemplateFlatPattern")))
                    ScannerFonctionsFeatureManager(pNoeud, ListeFonctions, NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);

                pNoeud = pNoeud.GetNext();
            }
        }

        /// <summary>
        /// Renvoi la liste des fonctions de la vue à partir du FeatureManager
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="TypeDeLaFonction"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<eFonction> ListListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            List<eFonction> pListeFonction = new List<eFonction>();

            ScannerFonctionsFeatureManager(NoeudVue(), pListeFonction, NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);

            return pListeFonction;
        }

        //public ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        //{
        //    Debug.Print(MethodBase.GetCurrentMethod());

        //    List<eFonction> pListeFonctions = ListListeDesFonctionsDeArbre(NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);
        //    ArrayList pArrayFonctions = new ArrayList();

        //    if (pListeFonctions.Count > 0)
        //        pArrayFonctions = new ArrayList(pListeFonctions);

        //    return pArrayFonctions;
        //}

        /// <summary>
        /// Scanner le FeatureManager pour rechercher la vue
        /// </summary>
        /// <param name="Noeud"></param>
        /// <param name="NoeudVue"></param>
        private void ScannerVueFeatureManager(TreeControlItem Noeud, ref TreeControlItem NoeudVue)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            TreeControlItem pNoeud = Noeud.GetFirstChild();

            while (pNoeud != null)
            {

                if (pNoeud.Text == _SwVue.GetName2())
                {
                    NoeudVue = pNoeud;
                    return;
                }

                if (pNoeud.ObjectType != (int)swTreeControlItemType_e.swFeatureManagerItem_Component)
                    ScannerVueFeatureManager(pNoeud, ref NoeudVue);

                if (NoeudVue != null)
                    return;

                pNoeud = pNoeud.GetNext();
            }
        }

        /// <summary>
        /// Renvoi le noeud de la vue
        /// </summary>
        /// <returns></returns>
        private TreeControlItem NoeudVue()
        {
            TreeControlItem pNoeudVue = null;

            ScannerVueFeatureManager(Feuille.Dessin.Modele.SwModele.FeatureManager.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneTop), ref pNoeudVue);

            if (pNoeudVue != null)
                pNoeudVue = pNoeudVue.GetFirstChild();

            return pNoeudVue;
        }

        #endregion

    }
}
