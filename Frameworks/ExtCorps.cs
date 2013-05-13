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
    public interface IExtCorps
    {
        Body2 SwCorps { get; }
        ExtPiece Piece { get; }
        String Nom { get; set; }
        TypeCorps_e TypeDeCorps { get; }
        ExtDossier Dossier { get; }
        ExtFonction PremiereFonction { get; }
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
        int NbIntersection(ExtCorps Corps);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DF347C75-F3B1-43AE-B7C4-393811BEBCB4")]
    [ProgId("Frameworks.ExtCorps")]
    public class ExtCorps : IExtCorps, IComparable<ExtCorps>, IComparer<ExtCorps>, IEquatable<ExtCorps>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private Body2 _SwCorps;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtCorps() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Body2 associé.
        /// </summary>
        public Body2 SwCorps { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwCorps; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ExtPiece Piece { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Piece; } }

        /// <summary>
        /// Retourne ou défini le nom du corps.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return SwCorps.Name; } set { Debug.Info(MethodBase.GetCurrentMethod());  SwCorps.Name = value; } }

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
                            return TypeCorps_e.cProfil;
                    }
                }

                return TypeCorps_e.cAutre;
            }
        }

        /// <summary>
        /// Retourne le parent ExtDossier.
        /// </summary>
        public ExtDossier Dossier
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                foreach (ExtDossier pDossier in _Piece.ListeDesDossiersDePiecesSoudees(TypeDeCorps, true))
                {
                    foreach (ExtCorps pCorps in pDossier.ListListeDesCorps())
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
        public ExtFonction PremiereFonction
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtFonction pFonction = new ExtFonction();

                if (pFonction.Init(SwCorps.GetFeatures()[0], _Piece.Modele))
                    return pFonction;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtCorps.
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
        internal Boolean Init(Body2 SwCorps, ExtPiece Piece)
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
        /// Renvoi la liste des fonctions d'un corps filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<ExtFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = new List<ExtFonction>();
            
            foreach(Feature pSwFonction in _SwCorps.GetFeatures())
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pFonction.Init(pSwFonction, _Piece.Modele))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        ExtFonction pSousFonction = new ExtFonction();

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

            List<ExtFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, AvecLesSousFonctions);
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
        public int NbIntersection(ExtCorps Corps)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            MathTransform XFormBase = _Piece.Modele.Composant.SwComposant.Transform2;
            MathTransform XFormTest = Corps._Piece.Modele.Composant.SwComposant.Transform2;

            Body2 CopieCorpsBase = _SwCorps.Copy();
            CopieCorpsBase.ApplyTransform(XFormBase);

            Body2 CopieCorpsTest = Corps._SwCorps.Copy();
            CopieCorpsTest.ApplyTransform(XFormTest);

            // SWBODYINTERSECT = 15901
            int Err;
            object[] ListeCorpsIntersection = CopieCorpsBase.Operations2(15901, CopieCorpsTest, out Err);

            if (ListeCorpsIntersection == null)
                return 0;

            return ListeCorpsIntersection.GetLength(0);
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtCorps>.CompareTo(ExtCorps Corps)
        {
            String Nom1 =  _Piece.Modele.SwModele.GetPathName() + _SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtCorps>.Compare(ExtCorps Corps1, ExtCorps Corps2)
        {
            String Nom1 = Corps1.Piece.Modele.SwModele.GetPathName() + Corps1.Nom;
            String Nom2 = Corps2.Piece.Modele.SwModele.GetPathName() + Corps2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtCorps>.Equals(ExtCorps Corps)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + _SwCorps.Name;
            String Nom2 = Corps.Piece.Modele.SwModele.GetPathName() + Corps.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}
