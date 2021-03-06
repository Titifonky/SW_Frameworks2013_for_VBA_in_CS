﻿using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("FC29A423-E1E0-40D4-BC58-3CB893ECF9AF")]
    public interface IeObjet
    {
        eModele Modele { get; }
        dynamic Objet { get; }
        swSelectType_e TypeDeObjet { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("54AC1164-1C5C-4CB1-B839-5256537D7355")]
    [ProgId("Frameworks.eObjet")]
    public class eObjet : IeObjet
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eObjet).Name;

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private dynamic _SwObjet = null;
        private swSelectType_e _TypeObjet = swSelectType_e.swSelNOTHING;

        #endregion

        #region "Constructeur\Destructeur"

        public eObjet() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le modele associé.
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Retourne l'objet associé.
        /// </summary>
        public dynamic Objet { get { Log.Methode(cNOMCLASSE); return _SwObjet; } }

        /// <summary>
        /// Retourne le type de l'objet.
        /// </summary>
        public swSelectType_e TypeDeObjet { get { Log.Methode(cNOMCLASSE); return _TypeObjet; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eObjet.
        /// </summary>
        /// <param name="Modele"></param>
        /// <param name="SwObjet"></param>
        /// <param name="TypeDeObjet"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele, dynamic SwObjet, swSelectType_e TypeDeObjet)
        {
            Log.Methode(cNOMCLASSE);

            if ((SwObjet != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _TypeObjet = TypeDeObjet;
                _SwObjet = SwObjet;
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
