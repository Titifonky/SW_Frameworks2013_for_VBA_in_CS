using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("413147BE-5B70-11E2-A5B0-4BF46188709B")]
    public interface IExtDossier
    {
        BodyFolder SwDossier { get; }
        ExtPiece Piece { get; }
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

        public BodyFolder SwDossier { get { return _swDossier; } }

        public ExtPiece Piece { get { return _Piece; } }

        public Boolean EstExclu { get; set; }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(BodyFolder Dossier, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Dossier != null) && (Piece != null) && Piece.EstInitialise)
            {

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _swDossier = Dossier;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        public Boolean Est(TypeCorps_e TypeDeCorps)
        {
            throw new NotImplementedException();
        }

        #endregion

    }
}
