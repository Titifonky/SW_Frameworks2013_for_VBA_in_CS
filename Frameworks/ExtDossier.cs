using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("413147BE-5B70-11E2-A5B0-4BF46188709B")]
    public interface IExtDossier
    {
        Boolean EstExclu { get; set; }
        Boolean Est(TypeCorps_e TypeDeCorps);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("4ABD2208-5B70-11E2-A0B2-51F46188709B")]
    [ProgId("Frameworks.ExtPiece")]
    public class ExtDossier : IExtDossier
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private BodyFolder _swDossier;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtDossier()
        {
        }

        #endregion

        #region "Propriétés"

        public Boolean EstExclu { get; set; }

        #endregion

        #region "Méthodes"

        internal ExtDossier Init(BodyFolder Dossier, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Dossier != null) && (Piece != null) && (Piece.Init() != null))
            {

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _swDossier = Dossier;
                _EstInitialise = true;
                return this;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            _EstInitialise = false;
            return null;
        }

        internal ExtDossier Init()
        {
            if (_EstInitialise)
                return this;
            else
                return null;
        }

        public Boolean Est(TypeCorps_e TypeDeCorps)
        {
            throw new NotImplementedException();
        }

        #endregion

    }
}
