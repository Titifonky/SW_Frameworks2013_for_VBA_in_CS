using System;

namespace Framework_SW2013
{
    /// <summary>
    /// Méthode d'extension
    /// Renvoi la chaine répété Nb fois.
    /// </summary>
    internal static class StringExtension
    {
        public static string Repeter(this String Chaine, int Nb)
        {
            if (Nb > 0)
                return string.Join(Chaine, new string[Nb]);
            else
                return "";
        }
    } 
}
