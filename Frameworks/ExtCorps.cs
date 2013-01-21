using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text.RegularExpressions;

/////////////////////////// Implementation terminée ///////////////////////////

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
        ExtFonction PremiereFonction();
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DF347C75-F3B1-43AE-B7C4-393811BEBCB4")]
    [ProgId("Frameworks.ExtCorps")]
    public class ExtCorps : IExtCorps, IComparable<ExtCorps>, IComparer<ExtCorps>, IEquatable<ExtCorps>
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private Body2 _SwCorps;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtCorps() { }

        #endregion

        #region "Propriétés"

        public Body2 SwCorps { get { return _SwCorps; } }

        public ExtPiece Piece { get { return _Piece; } }

        public String Nom { get { return SwCorps.Name; } set { SwCorps.Name = value; } }

        public TypeCorps_e TypeDeCorps
        {
            get
            {
                foreach (Feature Fonction in SwCorps.GetFeatures())
                {
                    switch (Fonction.GetTypeName2())
                    {
                        case "WeldMemberFeat":
                            return TypeCorps_e.cProfil;
                        case "FlatPattern":
                            return TypeCorps_e.cTole;
                    }
                }

                return TypeCorps_e.cAutre;
            }
        }

        public ExtDossier Dossier
        {
            get
            {
                foreach (ExtDossier pDossier in _Piece.ListeDesDossiers(TypeDeCorps, true))
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

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Body2 SwCorps, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwCorps != null) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwCorps = SwCorps;

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        public ExtFonction PremiereFonction()
        {
            ExtFonction pFonction = new ExtFonction();

            if (pFonction.Init(SwCorps.GetFeatures()[0], _Piece.Modele))
                return pFonction;

            return null;
        }

        /// <summary>
        /// Renvoi la liste des fonctions liées à un corps
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<ExtFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

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

        public ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, AvecLesSousFonctions);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
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
