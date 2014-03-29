using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("75D0F041-D524-454D-A6E8-7337B9588633")]
    public interface IeBarre
    {
        eCorps Corps { get; }
        String Profil { get; }
        Double Longueur { get; }
        Double Angle1 { get; }
        Double Angle2 { get; }
        Double Masse { get; }
        String Materiau { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C974BBDC-3C5A-495F-A532-2A61F7509025")]
    [ProgId("Frameworks.eBarre")]
    public class eBarre : IeBarre
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eBarre).Name;

        private Boolean _EstInitialise = false;

        private eCorps _Corps = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eBarre() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ePiece.
        /// </summary>
        public eCorps Corps { get { Log.Methode(cNOMCLASSE); return _Corps; } }

        public String Profil
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_NOM).Valeur;
            }
        }

        public Double Longueur
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_LONGUEUR).Valeur, pFormat);
            }
        }

        public Double Angle1
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_ANGLE1).Valeur, pFormat);
            }
        }

        public Double Angle2
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_ANGLE2).Valeur, pFormat);
            }
        }

        public Double Masse
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_MASSE).Valeur, pFormat);
            }
        }

        public String Materiau
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_MATERIAU).Valeur;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eBarre.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Log.Methode(cNOMCLASSE);

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cBarre))
            {
                _Corps = Corps;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion
    }
}
