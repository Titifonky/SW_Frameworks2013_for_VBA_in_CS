using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    internal sealed class Debug
    {
        private static Debug _instance = null;
        private static readonly object _Lock = new object();

        private ExtSldWorks _SW;
        private int _NbTabulations = 0;
        private String _Tab = "   ";
        private String _CheminFichierExecution;
        private String _CheminFichierDebug;

        #region "Propriétés"

        public static Debug Instance
        {
            get
            {
                lock (_Lock)
                {
                    if (_instance == null)
                    {
                        _instance = new Debug();
                    }
                }
                return _instance;
            }
        }

        public int NbTabulations
        {
            get { return _NbTabulations; }
            set { _NbTabulations = value; }
        }


        #endregion

        #region "Constructeur\Destructeur"

        public Debug()
        {
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ExtSldWorks SW)
        {
            if (SW != null)
            {
                _SW = SW;
                String pDossierMacros = _SW.swSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMacros);

                _CheminFichierExecution = Path.Combine(pDossierMacros, "Execution.txt");
                _CheminFichierDebug = Path.Combine(pDossierMacros, "Debug.txt");
                StreamWriter pFichierExecution = new StreamWriter(_CheminFichierExecution, false, System.Text.Encoding.Unicode);
                StreamWriter pFichierDebug = new StreamWriter(_CheminFichierDebug, false, System.Text.Encoding.Unicode);
                pFichierExecution.Close();
                pFichierDebug.Close();

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

        public void DebugAjouterLigne(String Ligne)
        {
            StreamWriter pFichierDebug = new StreamWriter(_CheminFichierDebug, true, System.Text.Encoding.Unicode);
            pFichierDebug.WriteLine(Ligne);
            pFichierDebug.Close();
        }

        #endregion
    }
}
