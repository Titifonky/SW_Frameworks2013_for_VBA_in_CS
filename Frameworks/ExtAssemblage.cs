using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("CACBFAD0-5820-11E2-B60C-A9046188709B")]
    public interface IExtAssemblage
    {
        AssemblyDoc SwAssemblage { get; }
        ExtModele Modele { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("CF8CC568-5820-11E2-B525-AA046188709B")]
    [ProgId("Frameworks.ExtAssemblage")]
    public class ExtAssemblage : IExtAssemblage
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private AssemblyDoc _SwAssemblage;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtAssemblage() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet AssemblyDoc associé.
        /// </summary>
        public AssemblyDoc SwAssemblage { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwAssemblage; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtAssemblage.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtAssemblage.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(ExtModele Modele)
        {
            
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                Debug.Info(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwAssemblage = Modele.SwModele as AssemblyDoc;
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
