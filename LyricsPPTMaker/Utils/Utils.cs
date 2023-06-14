using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LyricsPPTMaker
{
    public static class Utils
    {
        public static float CMToPoint(float cm)
        {
            return (cm / 2.54f) * 72;
        }
        public static float PointToCM(float pt)
        {
            return (pt / 72) * 2.54f;
        }
        public static double CMToPixel(double cm)
        {
            return cm * (96 / 2.54);
        }
        public static double PixelToCM(double px)
        {
            return px / (96 / 2.54);
        }



        /// <summary>
        /// 다가오는 일요일의 날짜를 구한다.
        /// </summary>
        /// <returns>해당 날짜를 나타내는 DateTime Object</returns>
        public static DateTime GetComingSundayDate()
        {
            int daysRemain = (7 - (int)DateTime.Now.DayOfWeek) % 7;
            DateTime dt = DateTime.Now.AddDays(daysRemain);
            return dt;
        }

        public static bool IsLastSundayOfMonth()
        {
            DateTime dt = GetComingSundayDate();
            var next = dt.AddDays(7);
            if (next.Month != dt.Month) return true;
            else return false;
        }
    }
}
