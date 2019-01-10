using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GeoSharp2018.UtilClass
{
    public class ExtentCoord
    {

        private double lx;
        public double Lx
        {
            get { return lx; }
            set { lx = value; }
        }
        private double ly;
        public double Ly
        {
            get { return ly; }
            set { ly = value; }
        }
        private double rx;
        public double Rx
        {
            get { return rx; }
            set { rx = value; }
        }
        private double ry;
        public double Ry
        {
            get { return ry; }
            set { ry = value; }
        }

        public ExtentCoord()
        {

        }

        public ExtentCoord(double lx, double ly, double rx, double ry)
        {
            this.lx = lx;
            this.ly = ly;
            this.rx = rx;
            this.ry = ry;
        }

        public void SetCoor(double lx, double ly, double rx, double ry)
        {
            this.lx = lx;
            this.ly = ly;
            this.rx = rx;
            this.ry = ry;
        }

    }
}
