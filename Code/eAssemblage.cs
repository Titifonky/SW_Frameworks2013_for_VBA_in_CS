using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;

namespace Framework
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

        private static readonly String cNOMCLASSE = typeof(eAssemblage).Name;

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
        public AssemblyDoc SwAssemblage { get { Log.Methode(cNOMCLASSE); return _SwAssemblage; } }

        /// <summary>
        /// Retourne le parent eModele.
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eAssemblage.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eAssemblage.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {

            Log.Methode(cNOMCLASSE);

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                Log.Message(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwAssemblage = Modele.SwModele as AssemblyDoc;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
