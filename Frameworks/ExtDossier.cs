using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

/////////////////////////// Implementation terminée ///////////////////////////

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
    public class ExtDossier : IExtDossier, IComparable<ExtDossier>, IComparer<ExtDossier>, IEquatable<ExtDossier>
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private BodyFolder _SwDossier;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtDossier() { }

        #endregion

        #region "Propriétés"

        public BodyFolder SwDossier { get { return _SwDossier; } }

        public ExtPiece Piece { get { return _Piece; } }

        public String Nom { get { return SwDossier.GetFeature().Name; } set { SwDossier.GetFeature().Name = value; } }

        public Boolean EstExclu
        {
            get { return Convert.ToBoolean(SwDossier.GetFeature().ExcludeFromCutList); }
            set { SwDossier.GetFeature().ExcludeFromCutList = value; }
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
                if (pGestProps.Init(SwDossier.GetFeature().CustomPropertyManager, _Piece.Modele))
                    return pGestProps;

                return null;
            }
        }

        public ExtCorps PremierCorps
        {
            get
            {
                ExtCorps pCorps = new ExtCorps();
                if ((SwDossier.GetBodyCount() > 0) && pCorps.Init(SwDossier.GetBodies()[0], _Piece))
                    return pCorps;

                return null;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(BodyFolder SwDossier, ExtPiece Piece)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            if ((SwDossier != null) && (SwDossier.GetBodyCount() > 0) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwDossier = SwDossier;

                _Debug.DebugAjouterLigne("\t -> " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne("\t !!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        internal List<ExtCorps> ListListeDesCorps(String NomARechercher = "")
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            List<ExtCorps> pListeCorps = new List<ExtCorps>();

            foreach (Body2 pSwCorps in SwDossier.GetBodies())
            {
                if (Regex.IsMatch(pSwCorps.Name, NomARechercher))
                {
                    ExtCorps pCorps = new ExtCorps();
                    if (pCorps.Init(pSwCorps, _Piece))
                        pListeCorps.Add(pCorps);
                }
            }

            return pListeCorps;
        }

        public ArrayList ListeDesCorps(String NomARechercher = "")
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            List<ExtCorps> pListeCorps = ListListeDesCorps(NomARechercher);
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtDossier>.CompareTo(ExtDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtDossier>.Compare(ExtDossier Dossier1, ExtDossier Dossier2)
        {
            String Nom1 = Dossier1.Piece.Modele.SwModele.GetPathName() + Dossier1.Nom;
            String Nom2 = Dossier2.Piece.Modele.SwModele.GetPathName() + Dossier2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtDossier>.Equals(ExtDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion

    }
}
