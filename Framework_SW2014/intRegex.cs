using System;
using System.Text.RegularExpressions;

namespace Framework
{
    /// <summary>
    /// Méthode d'extension
    /// Renvoi la chaine répété Nb fois.
    /// </summary>
    internal static class exRegex
    {
        public static Boolean IsMatch(String Chaine, String Modele)
        {
            if (Modele == "")
                return true;

            return Regex.IsMatch(Chaine, Modele);
        }
    } 
}
