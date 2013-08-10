using System;
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
    [Guid("2C02C4E9-0F4C-4A33-B9E6-0141641A5FE5")]
    public interface IeCorps
    {
        Body2 SwCorps { get; }
        ePiece Piece { get; }
        eTole Tole { get; }
        eBarre Barre { get; }
        String Nom { get; set; }
        TypeCorps_e TypeDeCorps { get; }
        String Materiau { get; set; }
        Boolean Visible { get; set; }
        eDossier Dossier { get; }
        eFonction PremiereFonction { get; }
        void Selectionner(Boolean Ajouter = true);
        ArrayList ListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false);
        int NbIntersection(eCorps Corps);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DF347C75-F3B1-43AE-B7C4-393811BEBCB4")]
    [ProgId("Frameworks.eCorps")]
    public class eCorps : IeCorps, IComparable<eCorps>, IComparer<eCorps>, IEquatable<eCorps>
    {
#region "Variables locales"

        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private Body2 _SwCorps = null;
        private eTole _Tole = null;
        private eBarre _Barre = null;
        private String _Nom = "";
        private Object _PID = null;

#endregion

#region "Constructeur\Destructeur"

        public eCorps() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne l'objet Body2 associé.
        /// </summary>
        public Body2 SwCorps
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                // On recupère le corps avec le PID, comme ça, ya plus de pb.
                if (_PID != null) // && (_SwCorps == null))
                {
                    int pErreur = 0;
                    Body2 pSwCorps = Piece.Modele.SwModele.Extension.GetObjectByPersistReference3(_PID, out pErreur);
                    Debug.Print("PID Erreur : " + pErreur);
                    if ((pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok) || (pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Suppressed))
                        _SwCorps = pSwCorps;
                }

                return _SwCorps;
            }
        }

        /// <summary>
        /// Retourne le parent ePiece.
        /// </summary>
        public ePiece Piece { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Piece; } }

        /// <summary>
        /// Retourne l'objet Tole
        /// </summary>
        public eTole Tole
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Tole == null)
                {
                    _Tole = new eTole();
                    _Tole.Init(this);
                }

                if (_Tole.EstInitialise)
                    return _Tole;

                return null;
            }
        }

        /// <summary>
        /// Retourne l'objet Barre
        /// </summary>
        public eBarre Barre
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Barre == null)
                {
                    _Barre = new eBarre();
                    _Barre.Init(this);
                }

                if (_Barre.EstInitialise)
                    return _Barre;

                return null;
            }
        }

        /// <summary>
        /// Retourne ou défini le nom du corps.
        /// </summary>
        public String Nom
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return SwCorps.Name;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                FeatureManager SwGestFonc = _Piece.Modele.SwModele.FeatureManager;
                String pNom = value;
                int Indice = 1;
                while (SwGestFonc.IsNameUsed((int)swNameType_e.swBodyName, pNom))
                {
                    pNom += "_" + Indice;
                    Indice++;
                }

                SwCorps.Name = pNom;
                _Nom = SwCorps.Name;
            }
        }

        /// <summary>
        /// Retourne le type du corps.
        /// </summary>
        public TypeCorps_e TypeDeCorps
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                foreach (Feature Fonction in SwCorps.GetFeatures())
                {
                    switch (Fonction.GetTypeName2())
                    {
                        case "SheetMetal":
                        case "SMBaseFlange":
                        case "SolidToSheetMetal":
                        case "FlatPattern":
                            return TypeCorps_e.cTole;
                        case "WeldMemberFeat":
                            return TypeCorps_e.cBarre;
                    }
                }

                return TypeCorps_e.cAutre;
            }
        }

        public String Materiau
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                String Db = "";
                String pNomConfigActive = Piece.Modele.GestDeConfigurations.ConfigurationActive.Nom;
                String Materiau = SwCorps.GetMaterialPropertyName(pNomConfigActive, out Db);
                if (String.IsNullOrEmpty(Materiau))
                    Materiau = Piece.Materiau;

                return Materiau;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                String[] pBaseDeDonnees = Piece.Modele.SW.SwSW.GetMaterialDatabases();

                String pNomConfigActive = Piece.Modele.GestDeConfigurations.ConfigurationActive.Nom;
                Selectionner(false);
                // On test si pour chaque Base de donnée si le matériau à bien été appliqué.
                // Si oui, on sort de la boucle
                foreach (String Bdd in pBaseDeDonnees)
                {
                    if (SwCorps.SetMaterialProperty(pNomConfigActive, Bdd, value) == (int)swBodyMaterialApplicationError_e.swBodyMaterialApplicationError_NoError)
                        break;
                    //Piece.SwPiece.SetMaterialPropertyName2(pNomConfigActive, Bdd, value);
                    //if (Materiau == value)
                    //    break;
                }
            }
        }

        public Boolean Visible
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return !SwCorps.DisableDisplay;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                SwCorps.DisableDisplay = !value;
                SwCorps.HideBody(!value);
                Selectionner(false);
                if (value)
                    Piece.Modele.SwModele.FeatureManager.ShowBodies();
                else
                    Piece.Modele.SwModele.FeatureManager.HideBodies();
                Piece.Modele.EffacerLesSelections();
            }
        }

        /// <summary>
        /// Retourne le parent eDossier.
        /// </summary>
        public eDossier Dossier
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                
                String pNom = Nom;
                
                foreach (eDossier pDossier in _Piece.ListeDesDossiersDePiecesSoudees(TypeDeCorps, true))
                {
                    List<eCorps> pListeCorps = pDossier.ListListeDesCorps("^" + Regex.Escape(pNom) + "$");
                    if (pListeCorps.Count > 0)
                        return pDossier;
                }

                return null;
            }
        }

        /// <summary>
        /// Renvoi la première fonction du corps.
        /// </summary>
        public eFonction PremiereFonction
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eFonction pFonction = new eFonction();

                if (pFonction.Init(SwCorps.GetFeatures()[0], _Piece.Modele))
                    return pFonction;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eBarre.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Piece"></param>
        /// <returns></returns>
        internal Boolean Init(Body2 SwCorps, ePiece Piece)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwCorps != null) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwCorps = SwCorps;
                _Nom = SwCorps.Name;
                _PID = Piece.Modele.SwModele.Extension.GetPersistReference3(_SwCorps);

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
        /// Initialiser l'objet eCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Body2 SwCorps, eModele Modele)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            Init(SwCorps, Modele.Piece);
            return _EstInitialise;
        }

        /// <summary>
        /// Selectionner le corps
        /// </summary>
        /// <param name="Ajouter"></param>
        public void Selectionner(Boolean Ajouter = true)
        {
            //SelectionMgr pSwSelMgr = Piece.Modele.SwModele.SelectionManager;
            //SelectData pSelData = default(SelectData);
            //pSelData = pSwSelMgr.CreateSelectData();
            //pSelData.Mark = -1;
            SwCorps.Select2(Ajouter, null);
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des fonctions d'un corps filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<eFonction> ListListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = new List<eFonction>();

            foreach (Feature pSwFonction in SwCorps.GetFeatures())
            {
                eFonction pFonction = new eFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher))
                    && (Regex.IsMatch(pSwFonction.GetTypeName2(), TypeDeLaFonction))
                    && pFonction.Init(pSwFonction, _Piece.Modele))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        eFonction pSousFonction = new eFonction();

                        if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher))
                            && (Regex.IsMatch(pSwSousFonction.GetTypeName2(), TypeDeLaFonction))
                            && pSousFonction.Init(pSwSousFonction, _Piece.Modele))
                            pListeFonctions.Add(pSousFonction);

                        pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                    }
                }
            }

            return pListeFonctions;
        }

        /// <summary>
        /// Renvoi la liste des fonctions d'un corps filtrée par les arguments.
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
        /// Renvoi le nb d'intersection avec le corps
        /// </summary>
        /// <param name="Composant"></param>
        /// <returns></returns>
        public int NbIntersection(eCorps Corps)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            int pNbInt = 0;

            MathTransform pXFormBase = _Piece.Modele.Composant.SwComposant.Transform2;
            MathTransform pXFormTest = Corps.Piece.Modele.Composant.SwComposant.Transform2;

            Body2 pCopieCorpsBase = SwCorps.Copy();
            Body2 pCopieCorpsTest = Corps.SwCorps.Copy();

            Debug.Print(pCopieCorpsBase.ApplyTransform(pXFormBase).ToString());
            Debug.Print(pCopieCorpsTest.ApplyTransform(pXFormTest).ToString());

            // SWBODYINTERSECT = 15901
            int Err;
            Object[] pListeCorpsIntersection = pCopieCorpsBase.Operations2(15901, pCopieCorpsTest, out Err);

            if (pListeCorpsIntersection != null)
                pNbInt = pListeCorpsIntersection.GetLength(0);

            pListeCorpsIntersection = null;

            pCopieCorpsBase = null;
            pCopieCorpsTest = null;
            pXFormBase = null;
            pXFormTest = null;

            return pNbInt;
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(eCorps Corps)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public int Compare(eCorps Corps1, eCorps Corps2)
        {
            String Nom1 = Corps1.Piece.Modele.SwModele.GetPathName() + Corps1.Nom;
            String Nom2 = Corps2.Piece.Modele.SwModele.GetPathName() + Corps2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public Boolean Equals(eCorps Corps)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.Equals(Nom2);
        }

#endregion
    }
}
