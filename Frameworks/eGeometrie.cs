using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17FF6AA9-3961-49A1-A800-CBED2CE3EF50")]
    public interface IePoint
    {
        Double X { get; set; }
        Double Y { get; set; }
        Double Z { get; set; }
        void Deplacer(eVecteur V);
        void Echelle(Double S);
        ePoint Additionner(eVecteur V);
        ePoint Multiplier(Double S);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BE5AB936-269C-45E6-8F8E-3496C744D044")]
    [ProgId("Frameworks.ePoint")]
    public class ePoint : IePoint
    {
#region "Variables locales"

#endregion

#region "Constructeur\Destructeur"

        public ePoint() { }

        public ePoint(Double X, Double Y, Double Z) { this.X = X; this.Y = Y; this.Z = Z; }

#endregion

#region "Propriétés"

        public Double X { get; set; }
        public Double Y { get; set; }
        public Double Z { get; set; }

#endregion

#region "Méthodes"

        public void Deplacer(eVecteur V)
        {
            X += V.X; Y += V.Y; Z += V.Z;
        }

        public void Echelle(Double S)
        {
            X *= S; Y *= S; Z *= S;
        }

        public ePoint Additionner(eVecteur V)
        {
            return new ePoint(X + V.X, Y + V.Y, Z + V.Z);
        }

        public ePoint Multiplier(Double S)
        {
            return new ePoint(X * S, Y * S, Z * S);
        }

#endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("F2D37BBD-3BB2-43C2-BD47-2443671839B4")]
    public interface IeVecteur
    {
        Double X { get; set; }
        Double Y { get; set; }
        Double Z { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BCFA38D4-2C36-494C-81C8-857FFF4037EE")]
    [ProgId("Frameworks.eVecteur")]
    public class eVecteur : IeVecteur
    {
#region "Variables locales"

#endregion

#region "Constructeur\Destructeur"

        public eVecteur() { }

        public eVecteur(Double X, Double Y, Double Z) { this.X = X; this.Y = Y; this.Z = Z; }

#endregion

#region "Propriétés"

        public Double X { get; set; }
        public Double Y { get; set; }
        public Double Z { get; set; }

#endregion

#region "Méthodes"

#endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("1037F0BB-EB59-4C63-A190-E36C403335FA")]
    public interface IeRepere
    {
        ePoint Origine { get; set; }
        eVecteur VecteurX { get; set; }
        eVecteur VecteurY { get; set; }
        eVecteur VecteurZ { get; set; }
        Double Echelle { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD8CE9F7-6C34-4202-AFE5-6C023B8DF4F4")]
    [ProgId("Frameworks.eRepere")]
    public class eRepere : IeRepere
    {
#region "Variables locales"
        
        private ePoint _Origine = new ePoint(0, 0, 0);
        private eVecteur _VecteurX = new eVecteur(0, 0, 0);
        private eVecteur _VecteurY = new eVecteur(0, 0, 0);
        private eVecteur _VecteurZ = new eVecteur(0, 0, 0);

#endregion

#region "Constructeur\Destructeur"

        public eRepere() { }

#endregion

#region "Propriétés"

        public ePoint Origine { get { return _Origine; } set { _Origine = value; } }
        public eVecteur VecteurX { get { return _VecteurX; } set { _VecteurX = value; } }
        public eVecteur VecteurY { get { return _VecteurY; } set { _VecteurY = value; } }
        public eVecteur VecteurZ { get { return _VecteurZ; } set { _VecteurZ = value; } }
        public Double Echelle { get; set; }

#endregion

#region "Méthodes"

#endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("BE8C3118-E918-4A72-9540-26C78B8B3498")]
    public interface IeRectangle
    {
        Double Lg { get; set; }
        Double Ht { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("80D861E1-18EA-4AB0-A1FF-29BD825A2B6C")]
    [ProgId("Frameworks.eRectangle")]
    public class eRectangle : IeRectangle
    {
#region "Variables locales"

#endregion

#region "Constructeur\Destructeur"

        public eRectangle() { }

#endregion

#region "Propriétés"

        public Double Lg { get; set; }
        public Double Ht { get; set; }

#endregion

#region "Méthodes"

#endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("4CA9B9F7-C61E-4B38-B01D-FA8D15DBA7C5")]
    public interface IeZone
    {
        ePoint PointMin { get; set; }
        ePoint PointMax { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D739C7A3-0EB2-4AA3-933B-07EA41288558")]
    [ProgId("Frameworks.eZone")]
    public class eZone : IeZone
    {
#region "Variables locales"
        
        private ePoint _PointMin = new ePoint(0, 0, 0);
        private ePoint _PointMax = new ePoint(0, 0, 0);

#endregion

#region "Constructeur\Destructeur"

        public eZone() { }

#endregion

#region "Propriétés"

        public ePoint PointMin { get { return _PointMin; } set { _PointMin = value; } }
        public ePoint PointMax { get { return _PointMax; } set { _PointMax = value; } }

#endregion

#region "Méthodes"

#endregion
    }

}
