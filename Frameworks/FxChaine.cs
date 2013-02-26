using System;

namespace Framework_SW2013
{
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
