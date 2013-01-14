using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("413147BE-5B70-11E2-A5B0-4BF46188709B")]
    public interface IExtDossier
    {
        BodyFolder SwDossier { get; }
        ExtPiece Piece { get; }
        String Nom { get; set; }
        Boolean EstExclu { get; set; }
        TypeCorps_e TypeDeCorps { get; }
        GestDeProprietes GestDeProprietes { get; }
        ExtCorps PremierCorps { get; }
        ArrayList ListeDesCorps(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("4ABD2208-5B70-11E2-A0B2-51F46188709B")]
    [ProgId("Frameworks.ExtPiece")]
    public class ExtDossier : IExtDossier
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private BodyFolder _swDossier;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtDossier()
        {
        }

        #endregion

        #region "Propriétés"

        public BodyFolder SwDossier { get { return _swDossier; } }

        public ExtPiece Piece { get { return _Piece; } }

        public String Nom { get { return _swDossier.GetFeature().Name; } set { _swDossier.GetFeature().Name = value; } }

        public Boolean EstExclu
        {
            get
            {
                if (_swDossier.GetFeature().ExcludeFromCutList == false)
                    return false;
                else
                    return true;
            }
            set
            {
                _swDossier.GetFeature().ExcludeFromCutList = value;
            }
        }

        public TypeCorps_e TypeDeCorps
        {
            get
            {
                ExtCorps pCorps = PremierCorps;
                if (pCorps.EstInitialise)
                    return pCorps.TypeDeCorps;

                return TypeCorps_e.cAucun;
            }
        }

        public GestDeProprietes GestDeProprietes
        {
            get
            {
                GestDeProprietes pGestProps = new GestDeProprietes();
                if (pGestProps.Init(_swDossier.GetFeature().CustomPropertyManager, _Piece.Modele))
                    return pGestProps;

                return null;
            }
        }

        ExtCorps PremierCorps
        {
            get
            {
                ExtCorps pCorps = new ExtCorps();
                if ((_swDossier.GetBodyCount() > 0) && pCorps.Init(_swDossier.GetBodies()[0], _Piece))
                    return pCorps;

                return null;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(BodyFolder SwDossier, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwDossier != null) && (SwDossier.GetBodyCount() > 0)  && (Piece != null) && Piece.EstInitialise)
            {

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _swDossier = SwDossier;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        internal List<ExtCorps> ListListeDesCorps(String NomARechercher = "")
        {
            List<ExtCorps> pListeCorps = new List<ExtCorps>();

            foreach (Body2 pSwCorps in _swDossier.GetBodies())
            {
                if (Regex.IsMatch(pSwCorps.Name,NomARechercher))
                {
                    ExtCorps Corps = new ExtCorps();
                    if (Corps.Init(pSwCorps, _Piece))
                        pListeCorps.Add(Corps);
                }
            }

            return pListeCorps;
        }

        public ArrayList ListeDesCorps(String NomARechercher = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtCorps> pListeCorps = ListListeDesCorps(NomARechercher);
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        #endregion

    }
}
