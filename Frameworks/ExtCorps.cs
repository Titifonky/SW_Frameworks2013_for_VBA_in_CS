using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;

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
    public class ExtCorps : IExtCorps
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private Body2 _swCorps;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtCorps() { }

        #endregion

        #region "Propriétés"

        public Body2 SwCorps { get { return _swCorps; } }

        public ExtPiece Piece { get { return _Piece; } }

        public String Nom { get { return _swCorps.Name; } set { _swCorps.Name = value; } }

        public TypeCorps_e TypeDeCorps
        {
            get
            {
                foreach (Feature Fonction in _swCorps.GetFeatures())
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

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _swCorps = SwCorps;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        ExtFonction PremiereFonction()
        {
            ExtFonction pFonction = new ExtFonction();

            if (pFonction.Init(_swCorps.GetFeatures()[0], _Piece))
                return pFonction;

            return null;
        }

        #endregion
    }
}
