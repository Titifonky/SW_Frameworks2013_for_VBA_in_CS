using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using System.Reflection;

namespace Framework_SW2013
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
        eCorps PremierCorps { get; }
        ArrayList ListeDesCorps(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("4ABD2208-5B70-11E2-A0B2-51F46188709B")]
    [ProgId("Frameworks.ePiece")]
    public class eDossier : IeDossier, IComparable<eDossier>, IComparer<eDossier>, IEquatable<eDossier>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private BodyFolder _SwDossier = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eDossier() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet BodyFolder associé.
        /// </summary>
        public BodyFolder SwDossier { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwDossier; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Piece; } }

        /// <summary>
        /// Retourne ou défini le nom du dossier.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod()); return SwDossier.GetFeature().Name; } set { Debug.Info(MethodBase.GetCurrentMethod()); SwDossier.GetFeature().Name = value; } }

        /// <summary>
        /// Retourne ou défini si le dossier est exclu de la nomenclature.
        /// </summary>
        public Boolean EstExclu
        {
            get { Debug.Info(MethodBase.GetCurrentMethod());  return Convert.ToBoolean(SwDossier.GetFeature().ExcludeFromCutList); }
            set { Debug.Info(MethodBase.GetCurrentMethod());  SwDossier.GetFeature().ExcludeFromCutList = value; }
        }

        /// <summary>
        /// Retourne le type de corps du dossier.
        /// </summary>
        public TypeCorps_e TypeDeCorps
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eCorps pCorps = PremierCorps;
                if (pCorps.EstInitialise)
                    return pCorps.TypeDeCorps;

                return TypeCorps_e.cAucun;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés du dossier.
        /// </summary>
        public eGestDeProprietes GestDeProprietes
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eGestDeProprietes pGestProps = new eGestDeProprietes();
                if (pGestProps.Init(SwDossier.GetFeature().CustomPropertyManager, _Piece.Modele))
                    return pGestProps;

                return null;
            }
        }

        /// <summary>
        /// Retourne le premier corps du dossier.
        /// </summary>
        public eCorps PremierCorps
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

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
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwDossier != null) && (SwDossier.GetBodyCount() > 0) && (Piece != null) && Piece.EstInitialise)
            {
                _Piece = Piece;
                _SwDossier = SwDossier;

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
        /// Méthode interne
        /// Renvoi la liste des corps du dossier filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<eCorps> ListListeDesCorps(String NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = new List<eCorps>();

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

        /// <summary>
        /// Renvoi la liste des corps du dossier filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesCorps(String NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = ListListeDesCorps(NomARechercher);
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<eDossier>.CompareTo(eDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<eDossier>.Compare(eDossier Dossier1, eDossier Dossier2)
        {
            String Nom1 = Dossier1.Piece.Modele.SwModele.GetPathName() + Dossier1.Nom;
            String Nom2 = Dossier2.Piece.Modele.SwModele.GetPathName() + Dossier2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<eDossier>.Equals(eDossier Dossier)
        {
            String Nom1 = _Piece.Modele.SwModele.GetPathName() + Nom;
            String Nom2 = Dossier.Piece.Modele.SwModele.GetPathName() + Dossier.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion

    }
}
