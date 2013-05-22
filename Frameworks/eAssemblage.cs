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
    public interface IeAssemblage
    {
        AssemblyDoc SwAssemblage { get; }
        eModele Modele { get; }
        void EditerLeComposant(eComposant Composant);
        void EditerAssemblage();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("CF8CC568-5820-11E2-B525-AA046188709B")]
    [ProgId("Frameworks.eAssemblage")]
    public class eAssemblage : IeAssemblage
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private AssemblyDoc _SwAssemblage = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eAssemblage() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet AssemblyDoc associé.
        /// </summary>
        public AssemblyDoc SwAssemblage { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwAssemblage; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

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
        internal Boolean Init(eModele Modele)
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

        /// <summary>
        /// Editer le composant dans le contexte de l'assemblage
        /// </summary>
        public void EditerLeComposant(eComposant Composant)
        {
            Modele.EffacerLesSelections();
            Composant.SwComposant.Select4(false, null, false);
            int info = 0;
            _SwAssemblage.EditPart2(true, false, ref info);
            Modele.EffacerLesSelections();
        }

        /// <summary>
        /// Edite l'assemblage racine
        /// </summary>
        public void EditerAssemblage()
        {
            _SwAssemblage.EditAssembly();
        }

        #endregion
    }
}
