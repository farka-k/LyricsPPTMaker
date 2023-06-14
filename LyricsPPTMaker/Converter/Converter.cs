using LyricsPPTMaker.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace LyricsPPTMaker
{
    public class FirstDegreeFunctionConverter : IValueConverter
    {
        public double A { get; set; }
        public double B { get; set; }

        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            double a = GetDoubleValue(parameter, A);

            double x = GetDoubleValue(value, 0.0);

            return (a * x) + B;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            double a = GetDoubleValue(parameter, A);

            double y = GetDoubleValue(value, 0.0);

            return (y - B) / a;
        }

        #endregion


        private double GetDoubleValue(object parameter, double defaultValue)
        {
            double a;
            if (parameter != null)
                try
                {
                    a = System.Convert.ToDouble(parameter);
                }
                catch
                {
                    a = defaultValue;
                }
            else
                a = defaultValue;
            return a;
        }
    }

    public class PositiveIntToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int retValue = (int)value + int.Parse((string)parameter);
            return (retValue > 0) ? retValue.ToString() : String.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class SongInfoItemConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values[1] == DependencyProperty.UnsetValue) return String.Empty;
            int index = (int)(values[1]);
            if (index < 0) return String.Empty;

            SongInfo item = ((ObservableCollection<SongInfo>)values[0])[index];
            if ((string)parameter == "Title")
            {
                return item.Title;
            }
            else if ((string)parameter == "Artist")
            {
                return item.Artist;
            }
            else if ((string)parameter == "Album")
            {
                return item.Album;
            }
            else
                return item.Lyrics;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class EnumtoBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (parameter.ToString() == value.ToString());
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? parameter : Binding.DoNothing;
        }
    }

    public class SlideSizeConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values[0] == DependencyProperty.UnsetValue) 
                return null;
            return (parameter.ToString() == ((SlideSizeType)values[0]).ToString());
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            object[] ret;
            if (parameter.ToString() == "WideScreen")
            {
                ret = new object[] { SlideSizeType.WideScreen, Constants.SlideSize16x9Width };
            }
            else
            {
                ret = new object[] { SlideSizeType.Normal, Constants.SlideSize4x3Width };
            }

            return (bool)value ? ret : new object[] { Binding.DoNothing, Binding.DoNothing };
        }
    }

    public class BoolToFontWeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? FontWeights.Bold : FontWeights.Normal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((FontWeight)value == FontWeights.Bold);
        }
    }

    public class BoolToFontStyleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? FontStyles.Italic : FontStyles.Normal;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((FontStyle)value == FontStyles.Italic);
        }
    }

    public class BoolToTextDecorationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? TextDecorations.Underline : null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return ((TextDecorationCollection)value == TextDecorations.Underline);
        }
    }

    public class LyricsVAlignmentConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((LyricsVAlignmentType)value == LyricsVAlignmentType.Top)
                return VerticalAlignment.Top;
            else if ((LyricsVAlignmentType)value == LyricsVAlignmentType.Center)
                return VerticalAlignment.Center;
            else
                return VerticalAlignment.Bottom;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class AlignmentOffsetToPaddingConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values[0] == DependencyProperty.UnsetValue)
                return null;
            LyricsVAlignmentType alignment = (LyricsVAlignmentType)(values[0]);
            double originalOffset = System.Convert.ToDouble(values[1]);
            double offset = Utils.CMToPixel(originalOffset + 0.7) / 4;
            Thickness ret;
            if (originalOffset < 0)
            {
                if (alignment == LyricsVAlignmentType.Center)
                    ret = new Thickness(0, 0, 0, -2 * offset);
                else
                    ret = new Thickness(0, 0, 0, -offset);
            }
            else
            {
                if (alignment == LyricsVAlignmentType.Center)
                    ret = new Thickness(0, 2 * offset, 0, 0);
                else
                    ret = new Thickness(0, offset, 0, 0);
            }
            return ret;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

}
