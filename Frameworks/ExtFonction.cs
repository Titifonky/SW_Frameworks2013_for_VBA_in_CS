using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0AE3E176-4DF1-40F2-92CD-9A7C83C8373A")]
    public interface IExtFonction
    {
        Body2 SwCorps { get; }
        ExtPiece Piece { get; }
        String Nom { get; set; }
        TypeCorps_e TypeDeCorps { get; }
        ExtFonction PremiereFonction();
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9009C0B9-61F1-42A1-AB7C-67DDF6AFB037")]
    [ProgId("Frameworks.ExtFonction")]
    public class ExtFonction : IExtFonction
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtPiece _Piece;
        private Feature _swFonction;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtFonction() { }

        #endregion

        #region "Propriétés"

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion
        #region "Méthodes"

        internal Boolean Init(Feature SwFonction, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwFonction != null) && (Piece != null) && Piece.EstInitialise)
            {

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _swFonction = SwFonction;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }
        #endregion
    }
}
