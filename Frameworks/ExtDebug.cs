using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("088A8454-50F6-11E2-AD1F-B6DF6188709B")]
    public interface IExtDebug
    {
        String Contenu { get; }
        int NbTabulations { get; set; }
        void AjouterLigne(String Ligne);
        void Effacer();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("10D61CCC-50F6-11E2-B9F2-BADF6188709B")]
    [ProgId("Frameworks.ExtDebug")]
    public class ExtDebug : IExtDebug
    {
        string _Texte = "";
        int _NbTabulations = 0;
        string _Tab = "   ";

        #region "Constructeur\Destructeur"

        public ExtDebug()
        {
        }

        #endregion

        #region "Propriétés"

        public int NbTabulations
        {
            get { return _NbTabulations; }
            set { _NbTabulations = value; }
        }

        public String Contenu
        {
            get { return _Texte; }
        }

        #endregion

        #region "Méthodes"

        public void AjouterLigne(String Ligne)
        {
            _Texte = _Texte + Environment.NewLine + _Tab.Repeter(_NbTabulations) + Ligne;
        }
        public void Effacer()
        {
            _Texte = "";
        }

        #endregion
    }
}
