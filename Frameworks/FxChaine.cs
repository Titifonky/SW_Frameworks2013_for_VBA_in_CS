
namespace Framework2013
{
    internal static class StringExtension
    {
        public static string Repeter(this string Chaine, int Nb)
        {
            return string.Join(Chaine, new string[Nb + 1]);
        }
    } 
}
