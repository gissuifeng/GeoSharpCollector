using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GeoSharp2018.UtilClass
{
    public class CoordinateTranslatorClass
    {

    }

    public class Gps
    {

        private double wgLat;
        private double wgLon;

        public Gps(double wgLat, double wgLon)
        {
            setWgLat(wgLat);
            setWgLon(wgLon);
        }

        public double getWgLat()
        {
            return wgLat;
        }

        public void setWgLat(double wgLat)
        {
            this.wgLat = wgLat;
        }

        public double getWgLon()
        {
            return wgLon;
        }

        public void setWgLon(double wgLon)
        {
            this.wgLon = wgLon;
        }

        public String toString()
        {
            return wgLat + "," + wgLon;
        }
    }

    public class CoordUtil
    {
        public static String BAIDU_LBS_TYPE = "bd09ll";

        public static double pi = 3.1415926535897932384626;
        public static double a = 6378245.0;
        public static double ee = 0.00669342162296594323;

        public static Gps gps84_To_Gcj02(double lat, double lon)
        {
           
            double dLat = transformLat(lon - 105.0, lat - 35.0);
            double dLon = transformLon(lon - 105.0, lat - 35.0);
            double radLat = lat / 180.0 * pi;
            double magic = Math.Sin(radLat);
            magic = 1 - ee * magic * magic;
            double sqrtMagic = Math.Sqrt(magic);
            dLat = (dLat * 180.0) / ((a * (1 - ee)) / (magic * sqrtMagic) * pi);
            dLon = (dLon * 180.0) / (a / sqrtMagic * Math.Cos(radLat) * pi);
            double mgLat = lat + dLat;
            double mgLon = lon + dLon;
            return new Gps(mgLat, mgLon);
        }

        public static Gps gcj_To_Gps84(double lat, double lon)
        {
            Gps gps = transform(lat, lon);
            double lontitude = lon * 2 - gps.getWgLon();
            double latitude = lat * 2 - gps.getWgLat();
            return new Gps(latitude, lontitude);
        }

        public static Gps gcj02_To_Bd09(double gg_lat, double gg_lon)
        {
            double x = gg_lon, y = gg_lat;
            double z = Math.Sqrt(x * x + y * y) + 0.00002 * Math.Sin(y * pi);
            double theta = Math.Atan2(y, x) + 0.000003 * Math.Cos(x * pi);
            double bd_lon = z * Math.Cos(theta) + 0.0065;
            double bd_lat = z * Math.Sin(theta) + 0.006;
            return new Gps(bd_lat, bd_lon);
        }

        public static Gps bd09_To_Gcj02(double bd_lat, double bd_lon)
        {
            double x = bd_lon - 0.0065, y = bd_lat - 0.006;
            double z = Math.Sqrt(x * x + y * y) - 0.00002 * Math.Sin(y * pi);
            double theta = Math.Atan2(y, x) - 0.000003 * Math.Cos(x * pi);
            double gg_lon = z * Math.Cos(theta);
            double gg_lat = z * Math.Sin(theta);
            return new Gps(gg_lat, gg_lon);
        }


        public static Gps bd09_To_Gps84(double bd_lat, double bd_lon)
        {

            Gps gcj02 = CoordUtil.bd09_To_Gcj02(bd_lat, bd_lon);
            Gps map84 = CoordUtil.gcj_To_Gps84(gcj02.getWgLat(),
                    gcj02.getWgLon());
            return map84;

        }

        public static Gps Gps84_To_bd09(double gps_lat, double gps_lon)
        {
            Gps gps = gps84_To_Gcj02(gps_lat, gps_lon);

            Gps gps1 = gcj02_To_Bd09(gps.getWgLat(), gps.getWgLon());

            return gps1;
        }

        public static bool outOfChina(double lat, double lon)
        {
            if (lon < 72.004 || lon > 137.8347)
                return true;
            if (lat < 0.8293 || lat > 55.8271)
                return true;
            return false;
        }

        public static Gps transform(double lat, double lon)
        {
            double dLat = transformLat(lon - 105.0, lat - 35.0);
            double dLon = transformLon(lon - 105.0, lat - 35.0);
            double radLat = lat / 180.0 * pi;
            double magic = Math.Sin(radLat);
            magic = 1 - ee * magic * magic;
            double sqrtMagic = Math.Sqrt(magic);
            dLat = (dLat * 180.0) / ((a * (1 - ee)) / (magic * sqrtMagic) * pi);
            dLon = (dLon * 180.0) / (a / sqrtMagic * Math.Cos(radLat) * pi);
            double mgLat = lat + dLat;
            double mgLon = lon + dLon;
            return new Gps(mgLat, mgLon);
        }

        public static double transformLat(double x, double y)
        {
            double ret = -100.0 + 2.0 * x + 3.0 * y + 0.2 * y * y + 0.1 * x * y
                    + 0.2 * Math.Sqrt(Math.Abs(x));
            ret += (20.0 * Math.Sin(6.0 * x * pi) + 20.0 * Math.Sin(2.0 * x * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(y * pi) + 40.0 * Math.Sin(y / 3.0 * pi)) * 2.0 / 3.0;
            ret += (160.0 * Math.Sin(y / 12.0 * pi) + 320 * Math.Sin(y * pi / 30.0)) * 2.0 / 3.0;
            return ret;
        }

        public static double transformLon(double x, double y)
        {
            double ret = 300.0 + x + 2.0 * y + 0.1 * x * x + 0.1 * x * y + 0.1
                    * Math.Sqrt(Math.Abs(x));
            ret += (20.0 * Math.Sin(6.0 * x * pi) + 20.0 * Math.Sin(2.0 * x * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(x * pi) + 40.0 * Math.Sin(x / 3.0 * pi)) * 2.0 / 3.0;
            ret += (150.0 * Math.Sin(x / 12.0 * pi) + 300.0 * Math.Sin(x / 30.0
                    * pi)) * 2.0 / 3.0;
            return ret;
        }
    }
}
