using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("413147BE-5B70-11E2-A5B0-4BF46188709B")]
    public interface IeDossier
    {
        BodyFolder SwDossier { get; }
        ePiece Piece { get; }
        String Nom { get; set; }
        Boolean EstExclu { get; set; }
        TypeCorps_e TypeDeCorps { get; }
        eGestDeProprietes GestDeProprietes { get; }
        int NbDeCorps { get; }
        eCorps PremierCorps { get; }
        ArrayList ListeDesCorps(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("4ABD2208-5B70-11E2-A0B2-51F46188709B")]
    [ProgId("Frameworks.ePiece")]
    public class eDossier : IeDossier, IComparable<eDossier>, IComparer<eDossier>, IEquatable<eDossier>
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eDossier).Name;

        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private BodyFolder _SwDossier = null;
        private Object _PID = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eDossier() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet BodyFolder associé.
        /// </summary>
        public BodyFolder SwDossier
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                if (_PID != null)
                {
                    int pErreur = 0;
                    Feature pSwFonction = _Piece.Modele.SwModele.Extension.GetObjectByPersistReference3(_PID, out pErreur);
                    Log.Message("PID Erreur : " + pErreur);
                    if ((pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok)
                        || (pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Suppressed))
                        _SwDossier = pSwFonction.GetSpecificFeature2();
                }
                else
                {
                    Log.Message("Pas de PID");
                    MajPID();
                }

                return _SwDossier;
            }
        }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Log.Methode(cNOMCLASSE); return _Piece; } }

        /// <summary>
        /// Retourne ou défini le nom du dossier.
        /// </summary>
        public String Nom { get { Log.Methode(cNOMCLASSE); return SwDossier.GetFeature().Name; } set { Log.Methode(cNOMCLASSE); SwDossier.GetFeature().Name = value; } }

        /// <summary>
        /// Retourne ou défini si le dossier est exclu de la nomenclature.
        /// </summary>
        public Boolean EstExclu
        {
            get { Log.Methode(cNOMCLASSE); return Convert.ToBoolean(SwDossier.GetFeature().ExcludeFromCutList); }
            set { Log.Methode(cNOMCLASSE); SwDossier.GetFeature().ExcludeFromCutList = value; }
        }

        /// <summary>
        /// Retourne le type de corps du dossier.
        /// </summary>
        public TypeCorps_e TypeDeCorps
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                if (EstUnDossierDeBarres(SwDossier))
                    return TypeCorps_e.cBarre;

                if (EstUnDossierDeToles(SwDossier))
                    return TypeCorps_e.cTole;

                eCorps pCorps = PremierCorps;
                if (pCorps.EstInitialise)
                    return pCorps.TypeDeCorps;

                return TypeCorps_e.cAucun;
            }
        }

        static internal Boolean EstUnDossierDeBarres(BodyFolder SwDossier)
        {
            CustomPropertyManager pPropMgr = SwDossier.GetFeature().CustomPropertyManager;

            foreach (String iNom in pPropMgr.GetNames())
            {
                String pVal, pResult;
                Boolean Resolved;

                pPropMgr.Get5(iNom, false, out pVal, out pResult, out Resolved);

                Log.Message("                        " + pVal);
                if (Regex.IsMatch(pVal, "^\"LENGTH@@@"))
                    return true;
            }

            return false;
        }

        static internal Boolean EstUnDossierDeToles(BodyFolder SwDossier)
        {
            CustomPropertyManager pPropMgr = SwDossier.GetFeature().CustomPropertyManager;

            foreach (String iNom in pPropMgr.GetNames())
            {
                String pVal, pResult;
                Boolean Resolved;

                pPropMgr.Get5(iNom, false, out pVal, out pResult, out Resolved);

                Log.Message("                        " + pVal);
                if (Regex.IsMatch(pVal, "^\"SW-Longueur du flanc de tôle@@@"))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés du dossier.
        /// </summary>
        public eGestDeProprietes GestDeProprietes
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                eGestDeProprietes pGestProps = new eGestDeProprietes();
                if (pGestProps.Init(SwDossier.GetFeature().CustomPropertyManager, _Piece.Modele))
                    return pGestProps;

                return null;
            }
        }

        public int NbDeCorps
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return SwDossier.GetBodyCount();
            }
        }

        /// <summary>
        /// Retourne le premier corps du dossier.
        /// </summary>
        public eCorps PremierCorps
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                eCorps pCorps = new eCorps();
                if ((SwDossier.GetBodyCount() > 0) && pCorps.Init(SwDossier.GetBodies()[0], _Piece))
                    return pCorps;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtDossier.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtDossier.
        /// </summary>
        /// <param name="SwDossier"></param>
        /// <param name="Piece"></param>
        /// <returns></returns>
        internal Boolean Init(BodyFolder SwDossier, ePiece Piece)
        {
            Log.Methode(cNOMCLASSE);

            if ((SwDossier != null) && (SwDossier.GetBodyCount() > 0) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwDossier = SwDossier;
                MajPID();
                Log.Message(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        private void MajPID()
        {
            Log.Methode(cNOMCLASSE);

            if (_SwDossier == null)
                return;
            _PID = _Piece.Modele.SwModele.Extension.GetPersistReference3(_SwDossier.GetFeature());
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des corps du dossier filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesCorps(String NomARechercher = "")
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pListeCorps = new ArrayList();

            foreach (Body2 pSwCorps in SwDossier.GetBodies())
            {
                if (Regex.IsMatch(pSwCorps.Name, NomARechercher))
                {
                    eCorps pCorps = new eCorps();
                    if (pCorps.Init(pSwCorps, _Piece))
                        pListeCorps.Add(pCorps);
                }
            }

            return pListeCorps;
        }

        #endregion

        #region "Interfaces génériques"

        public int CompareTo(eDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public int Compare(eDossier Dossier1, eDossier Dossier2)
        {
            String Nom1 = Dossier1.Piece.Modele.SwModele.GetPathName() + Dossier1.Nom;
            String Nom2 = Dossier2.Piece.Modele.SwModele.GetPathName() + Dossier2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public Boolean Equals(eDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion

    }
}
