using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
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
        Orientation_e Orientation { get; set; }
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

        private static readonly String cNOMCLASSE = typeof(eVue).Name;

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
        public View SwVue { get { Log.Methode(cNOMCLASSE); return _SwVue; } }

        /// <summary>
        /// Retourne le parent ExtFeuille.
        /// </summary>
        public eFeuille Feuille { get { Log.Methode(cNOMCLASSE); return _Feuille; } }

        /// <summary>
        /// Retourne ou défini le nom de la vue.
        /// </summary>
        public String Nom { get { Log.Methode(cNOMCLASSE); return _SwVue.GetName2(); } set { Log.Methode(cNOMCLASSE); _SwVue.SetName2(value); } }

        /// <summary>
        /// Retourne le modele ExtModele référencé par la vue.
        /// </summary>
        public eModele ModeleDeReference
        {
            get
            {
                Log.Methode(cNOMCLASSE);
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
                Log.Methode(cNOMCLASSE);
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
                Log.Methode(cNOMCLASSE);
                eDimensionVue pDimensions = new eDimensionVue();

                if (pDimensions.Init(this))
                    return pDimensions;

                return null;
            }
        }

        /// <summary>
        /// Orientation de la vue, Portrait ou Paysage
        /// </summary>
        public Orientation_e Orientation
        {
            get
            {
                eRectangle Rect = Dimensions.Rectangle;
                if (Rect.Ht > Rect.Lg)
                    return Orientation_e.cPortrait;

                return Orientation_e.cPaysage;
            }
            set
            {
                eDimensionVue Dim = Dimensions;
                if ((Orientation == Orientation_e.cPaysage) && (value == Orientation_e.cPortrait))
                    Dimensions.Angle = 90;
                else if ((Orientation == Orientation_e.cPortrait) && (value == Orientation_e.cPaysage))
                    Dimensions.Angle = -90;
            }
        }

        /// <summary>
        /// Afficher ou masquer les lignes de pliage
        /// </summary>
        public Boolean AfficherLignesDePliage
        {
            set
            {
                Log.Methode(cNOMCLASSE);

                ModelDoc2 pSwModele = Feuille.Dessin.Modele.SwModele;
                eConfiguration pConfig = ConfigurationDeReference;

                if (pConfig.Est(TypeConfig_e.cDepliee))
                {

                    eTole pTole = pConfig.ToleDeplie();

                    if (pTole != null)
                    {
                        eFonction FonctionLignePliage = pTole.FonctionDepliee.ListeDesSousFonctions(CONSTANTES.LIGNES_DE_PLIAGE)[0] as eFonction;
                        String NomFonctionLignePliagePourSelection = FonctionLignePliage.Nom + "@" + SwVue.RootDrawingComponent.Name + "@" + SwVue.Name;
                        Feuille.Dessin.Modele.SwModele.Extension.SelectByID2(NomFonctionLignePliagePourSelection, "SKETCH", 0, 0, 0, false, 0, null, 0);

                        if (value)
                            Feuille.Dessin.Modele.SwModele.UnblankSketch();
                        else
                            Feuille.Dessin.Modele.SwModele.BlankSketch();

                        Feuille.Dessin.Modele.EffacerLesSelections();
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
                Log.Methode(cNOMCLASSE);
                return _SwVue.ShowSheetMetalBendNotes;
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SwVue.ShowSheetMetalBendNotes = value;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((SwVue != null) && (Feuille != null) && Feuille.EstInitialise)
            {
                _Feuille = Feuille;
                _SwVue = SwVue;

                Log.Message(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
            Log.Methode(cNOMCLASSE);

            if ((SwVue != null) && (Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                eFeuille Feuille = new eFeuille();
                if (Feuille.Init(SwVue.Sheet, Modele.Dessin))
                {
                    _Feuille = Feuille;
                    _SwVue = SwVue;

                    Log.Message(this.Nom);
                    _EstInitialise = true;
                }
                else
                {
                    Log.Message("!!!!! Erreur d'initialisation");
                }
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Supprimer la vue
        /// </summary>
        public void Supprimer()
        {
            Log.Methode(cNOMCLASSE);

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
        private void ScannerFonctionsFeatureManager(TreeControlItem Noeud, ArrayList ListeFonctions, String NomARechercher, String TypeDeLaFonction, Boolean AvecLesSousFonctions)
        {
            Log.Methode(cNOMCLASSE);

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
        public ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            ArrayList pListeFonction = new ArrayList();

            ScannerFonctionsFeatureManager(NoeudVue(), pListeFonction, NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);

            return pListeFonction;
        }

        //public ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        //{
        //    Log.Print(cNOMCLASSE);

        //    List<eFonction> pListeFonctions = ListeDesFonctionsDeArbre(NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);
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
            Log.Methode(cNOMCLASSE);

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

            ScannerVueFeatureManager(Feuille.Dessin.Modele.GestDeFonction_NoeudRacine(), ref pNoeudVue);

            if (pNoeudVue != null)
                pNoeudVue = pNoeudVue.GetFirstChild();

            return pNoeudVue;
        }

        #endregion

    }
}
