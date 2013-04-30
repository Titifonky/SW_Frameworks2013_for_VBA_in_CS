using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17FF6AA9-3961-49A1-A800-CBED2CE3EF50")]
    public interface IPoint
    {
        Double X { get; set; }
        Double Y { get; set; }
        Double Z { get; set; }
        void Deplacer(Vecteur V);
        void Echelle(Double S);
        Point Additionner(Vecteur V);
        Point Multiplier(Double S);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BE5AB936-269C-45E6-8F8E-3496C744D044")]
    [ProgId("Frameworks.Point")]
    public class Point : IPoint
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public Point() { }

        public Point(Double X, Double Y, Double Z) { this.X = X; this.Y = Y; this.Z = Z; }

        #endregion

        #region "Propriétés"

        public Double X { get; set; }
        public Double Y { get; set; }
        public Double Z { get; set; }

        #endregion

        #region "Méthodes"

        public void Deplacer(Vecteur V)
        {
            X += V.X; Y += V.Y; Z += V.Z;
        }

        public void Echelle(Double S)
        {
            X *= S; Y *= S; Z *= S;
        }

        public Point Additionner(Vecteur V)
        {
            return new Point(X + V.X, Y + V.Y, Z + V.Z);
        }

        public Point Multiplier(Double S)
        {
            return new Point(X * S, Y * S, Z * S);
        }

        #endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("F2D37BBD-3BB2-43C2-BD47-2443671839B4")]
    public interface IVecteur
    {
        Double X { get; set; }
        Double Y { get; set; }
        Double Z { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BCFA38D4-2C36-494C-81C8-857FFF4037EE")]
    [ProgId("Frameworks.Vecteur")]
    public class Vecteur : IVecteur
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public Vecteur() { }

        public Vecteur(Double X, Double Y, Double Z) { this.X = X; this.Y = Y; this.Z = Z; }

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
    public interface IRepere
    {
        Point Origine { get; set; }
        Vecteur VecteurX { get; set; }
        Vecteur VecteurY { get; set; }
        Vecteur VecteurZ { get; set; }
        Double Echelle { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD8CE9F7-6C34-4202-AFE5-6C023B8DF4F4")]
    [ProgId("Frameworks.Repere")]
    public class Repere : IRepere
    {
        #region "Variables locales"
        
        private Point _Origine = new Point(0, 0, 0);
        private Vecteur _VecteurX = new Vecteur(0, 0, 0);
        private Vecteur _VecteurY = new Vecteur(0, 0, 0);
        private Vecteur _VecteurZ = new Vecteur(0, 0, 0);

        #endregion

        #region "Constructeur\Destructeur"

        public Repere() { }

        #endregion

        #region "Propriétés"

        public Point Origine { get { return _Origine; } set { _Origine = value; } }
        public Vecteur VecteurX { get { return _VecteurX; } set { _VecteurX = value; } }
        public Vecteur VecteurY { get { return _VecteurY; } set { _VecteurY = value; } }
        public Vecteur VecteurZ { get { return _VecteurZ; } set { _VecteurZ = value; } }
        public Double Echelle { get; set; }

        #endregion

        #region "Méthodes"

        #endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("BE8C3118-E918-4A72-9540-26C78B8B3498")]
    public interface IRectangle
    {
        Double Lg { get; set; }
        Double Ht { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("80D861E1-18EA-4AB0-A1FF-29BD825A2B6C")]
    [ProgId("Frameworks.Rectangle")]
    public class Rectangle : IRectangle
    {
        #region "Variables locales"

        #endregion

        #region "Constructeur\Destructeur"

        public Rectangle() { }

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
    public interface IZone
    {
        Point PointMin { get; set; }
        Point PointMax { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D739C7A3-0EB2-4AA3-933B-07EA41288558")]
    [ProgId("Frameworks.Zone")]
    public class Zone : IZone
    {
        #region "Variables locales"
        
        private Point _PointMin = new Point(0, 0, 0);
        private Point _PointMax = new Point(0, 0, 0);

        #endregion

        #region "Constructeur\Destructeur"

        public Zone() { }

        #endregion

        #region "Propriétés"

        public Point PointMin { get { return _PointMin; } set { _PointMin = value; } }
        public Point PointMax { get { return _PointMax; } set { _PointMax = value; } }

        #endregion

        #region "Méthodes"

        #endregion
    }

}
