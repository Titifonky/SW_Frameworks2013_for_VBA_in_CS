using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;

namespace Framework_SW2013
{
    internal static class Debug
    {
        #region "Propriétés"

        private static String _DOSSIER = Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
        private static String _FICHIER_BASE = Assembly.GetExecutingAssembly().FullName.Split(',')[0] + "_Debug";
        private static String _CHEMIN_FICHIER = "";
        private static int _NO_FICHIER = 1;
        private static long _NB_LIGNE_MAX = 50000;
        private static long _NB_LIGNE = 0;
        private static Boolean _ACTIF = false;
        private static Boolean _INIT = true;
        private static String[] LISTE_METHODE = new String[500];

        public static Boolean Actif { get { return _ACTIF; } set { _ACTIF = value; } }

        #endregion

        #region "Méthodes"

        public static void Init(SldWorks SldWks)
        {
            try
            {
                foreach (String F in Directory.GetFiles(_DOSSIER, _FICHIER_BASE + "*", SearchOption.TopDirectoryOnly))
                {
                    File.Delete(F);
                }

                _CHEMIN_FICHIER = Path.Combine(_DOSSIER, _FICHIER_BASE + "-" + _NO_FICHIER.ToString() + ".txt");
                StreamWriter pFichierDebug = new StreamWriter(_CHEMIN_FICHIER, false, System.Text.Encoding.Unicode);

                /// A chaque initialisation, on inscrit la version de SW
                String pVersionDeBase; String pVersionCourante; String pHotfixe;
                SldWks.GetBuildNumbers2(out pVersionDeBase, out pVersionCourante, out pHotfixe);
                String pRevision = SldWks.RevisionNumber();
                pFichierDebug.WriteLine("\n ");
                pFichierDebug.WriteLine("====================================================================================================");
                pFichierDebug.WriteLine("|                                                                                                  |");
                pFichierDebug.WriteLine("|                                       SOLIDWORKS DEBUG                                           |");
                pFichierDebug.WriteLine("|                                                                                                  |");
                pFichierDebug.WriteLine("====================================================================================================");
                pFichierDebug.WriteLine("Version de base : " + pVersionDeBase + "    Version courante : " + pVersionCourante + "    Hotfixe : " + pHotfixe + "    Revision : " + pRevision);
                pFichierDebug.WriteLine("----------------------------------------------------------------------------------------------------");
                pFichierDebug.WriteLine("\n ");

                pFichierDebug.Close();
            }
            catch
            {
                _INIT = false;
            }

        }

        /// <summary>
        /// Si Message est vide, écrit le nom de l'objet et de la méthode.
        /// Sinon, écrit Message.
        /// </summary>
        /// <param name="Message"></param>
        /// 
        public static void Info(String Message = "")
        {
            Info(null, Message);

        }

        public static void Info(_MethodBase Methode, String Message = "")
        {
            try
            {
                if (!(_ACTIF && _INIT))
                    return;

                StackTrace pPile = new StackTrace();

                int Pos = pPile.FrameCount - 1;
                Boolean AjouterAuFichier = true;

                if (_NB_LIGNE > _NB_LIGNE_MAX)
                {
                    AjouterAuFichier = false;
                    _NB_LIGNE = 0;
                    _NO_FICHIER++;
                    _CHEMIN_FICHIER = Path.Combine(_DOSSIER, _FICHIER_BASE + "-" + _NO_FICHIER.ToString() + ".txt");
                }

                if ((Methode != null) && (Pos < 500))
                {
                    LISTE_METHODE[Pos] = Methode.DeclaringType.ToString().Split('.')[1] + "." + Methode.Name;
                }

                if (!String.IsNullOrEmpty(Message))
                    Message = " -> " + Message;

                StreamWriter pFichierDebug = new StreamWriter(_CHEMIN_FICHIER, AjouterAuFichier, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(Pos) + LISTE_METHODE[Pos] + Message);
                pFichierDebug.Close();
                _NB_LIGNE++;
            }
            catch
            {
                StreamWriter pFichierDebug = new StreamWriter(_CHEMIN_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("Erreur");
                pFichierDebug.Close();
            }

        }

        #endregion
    }
}
