using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("B49F2CE6-5820-11E2-96DC-9A046188709B")]
    public interface IExtDebug
    {
        int NbTabulations { get; set; }
        Boolean Init(ExtSldWorks SW);
        void ExecutionAjouterLigne(String Ligne);
        void ErreurAjouterLigne(String Ligne);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("B91D00E0-5820-11E2-BECC-9E046188709B")]
    [ProgId("Frameworks.ExtDebug")]
    public class ExtDebug : IExtDebug
    {
        private ExtSldWorks _SW;
        private int _NbTabulations = 0;
        private String _Tab = "   ";
        private String _CheminFichierExecution;
        private String _CheminFichierErreur;

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


        #endregion

        #region "Méthodes"

        public Boolean Init(ExtSldWorks SW)
        {
            if (!(SW.Equals(null)))
            {
                _SW = SW;
                String pDossierMacros = _SW.swSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMacros);

                _CheminFichierExecution = Path.Combine(pDossierMacros, "Execution.txt");
                _CheminFichierErreur = Path.Combine(pDossierMacros, "Erreur.txt");
                StreamWriter pFichierExecution = new StreamWriter(_CheminFichierExecution, false, System.Text.Encoding.Unicode);
                StreamWriter pFichierErreur = new StreamWriter(_CheminFichierErreur, false, System.Text.Encoding.Unicode);
                pFichierExecution.Close();
                pFichierErreur.Close();

                return true;
            }

            return false;
        }

        public void ExecutionAjouterLigne(String Ligne)
        {
            StreamWriter pFichierExecution = new StreamWriter(_CheminFichierExecution, true, System.Text.Encoding.Unicode);
            pFichierExecution.WriteLine(_Tab.Repeter(_NbTabulations) + Ligne);
            pFichierExecution.Close();
        }

        public void ErreurAjouterLigne(String Ligne)
        {
            StreamWriter pFichierErreur = new StreamWriter(_CheminFichierErreur, true, System.Text.Encoding.Unicode);
            pFichierErreur.WriteLine(_Tab.Repeter(_NbTabulations) + Ligne);
            pFichierErreur.Close();
        }

        #endregion
    }
}
