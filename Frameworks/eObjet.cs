using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
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
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Retourne l'objet associé.
        /// </summary>
        public dynamic Objet { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwObjet; } }

        /// <summary>
        /// Retourne le type de l'objet.
        /// </summary>
        public swSelectType_e TypeDeObjet { get { Debug.Info(MethodBase.GetCurrentMethod()); return _TypeObjet; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwObjet != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _TypeObjet = TypeDeObjet;
                _SwObjet = SwObjet;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

#endregion
    }
}
