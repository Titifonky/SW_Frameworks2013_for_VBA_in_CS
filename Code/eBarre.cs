using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Reflection;
using System.Globalization;

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
        public eCorps Corps { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Corps; } }

        public String Profil
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_NOM).Valeur;
            }
        }

        public Double Longueur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_LONGUEUR).Valeur, pFormat);
            }
        }

        public Double Angle1
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_ANGLE1).Valeur, pFormat);
            }
        }

        public Double Angle2
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_ANGLE2).Valeur, pFormat);
            }
        }

        public Double Masse
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                NumberFormatInfo pFormat = new NumberFormatInfo();
                pFormat.NumberDecimalSeparator = ".";
                return Convert.ToDouble(_Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_MASSE).Valeur, pFormat);
            }
        }

        public String Materiau
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _Corps.Dossier.GestDeProprietes.RecupererPropriete(CONSTANTES.PROFIL_MATERIAU).Valeur;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cBarre))
            {
                _Corps = Corps;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

#endregion
    }
}
