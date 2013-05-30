using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;

namespace Framework_SW2013
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
        eDossier Dossier { get; }
        eFonction PremiereFonction { get; }
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
        
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

        #endregion

        #region "Constructeur\Destructeur"

        public eCorps() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Body2 associé.
        /// </summary>
        public Body2 SwCorps { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwCorps; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Piece; } }

        /// <summary>
        /// Retourne l'objet Tole
        /// </summary>
        public eTole Tole
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

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
                Debug.Info(MethodBase.GetCurrentMethod());

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
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwCorps.Name;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                FeatureManager SwGestFonc = _Piece.Modele.SwModele.FeatureManager;
                String pNom = value;
                int Indice = 1;
                while (SwGestFonc.IsNameUsed((int)swNameType_e.swBodyName, pNom))
                {
                    pNom += "_" + Indice;
                    Indice++;
                }

                SwCorps.Name = pNom;
            }
        }

        /// <summary>
        /// Retourne le type du corps.
        /// </summary>
        public TypeCorps_e TypeDeCorps
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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

        /// <summary>
        /// Retourne le parent ExtDossier.
        /// </summary>
        public eDossier Dossier
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                foreach (eDossier pDossier in _Piece.ListeDesDossiersDePiecesSoudees(TypeDeCorps, true))
                {
                    foreach (eCorps pCorps in pDossier.ListListeDesCorps())
                    {
                        if (pCorps.Nom == Nom)
                        {
                            return pDossier;
                        }
                    }
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
                Debug.Info(MethodBase.GetCurrentMethod());

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
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Piece"></param>
        /// <returns></returns>
        internal Boolean Init(Body2 SwCorps, ePiece Piece)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwCorps != null) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwCorps = SwCorps;

                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Body2 SwCorps, eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwCorps != null) && (Modele != null) && (Modele.TypeDuModele == TypeFichier_e.cPiece) && Modele.EstInitialise)
            {
                _Piece = Modele.Piece;
                _SwCorps = SwCorps;

                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des fonctions d'un corps filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<eFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = new List<eFonction>();
            
            foreach(Feature pSwFonction in _SwCorps.GetFeatures())
            {
                eFonction pFonction = new eFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pFonction.Init(pSwFonction, _Piece.Modele))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        eFonction pSousFonction = new eFonction();

                        if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pSousFonction.Init(pSwSousFonction, _Piece.Modele))
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
        public ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, AvecLesSousFonctions);
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
            Debug.Info(MethodBase.GetCurrentMethod());

            int pNbInt = 0;

            MathTransform pXFormBase = _Piece.Modele.Composant.SwComposant.Transform2;
            MathTransform pXFormTest = Corps.Piece.Modele.Composant.SwComposant.Transform2;

            Body2 pCopieCorpsBase = _SwCorps.Copy();
            Body2 pCopieCorpsTest = Corps._SwCorps.Copy();

            Debug.Info(pCopieCorpsBase.ApplyTransform(pXFormBase).ToString());
            Debug.Info(pCopieCorpsTest.ApplyTransform(pXFormTest).ToString());

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

        int IComparable<eCorps>.CompareTo(eCorps Corps)
        {
            String Nom1 =  _Piece.Modele.SwModele.GetPathName() + _SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<eCorps>.Compare(eCorps Corps1, eCorps Corps2)
        {
            String Nom1 = Corps1.Piece.Modele.SwModele.GetPathName() + Corps1.Nom;
            String Nom2 = Corps2.Piece.Modele.SwModele.GetPathName() + Corps2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<eCorps>.Equals(eCorps Corps)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + _SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.Equals(Nom2);
        }

        internal Boolean Equals(eCorps Corps)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + _SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}
