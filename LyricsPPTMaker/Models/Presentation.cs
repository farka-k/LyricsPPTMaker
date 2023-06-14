using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LyricsPPTMaker.Models
{
    public static class Constants
    {
        public static string BaseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        public static string ResourceDirectory = AppDomain.CurrentDomain.BaseDirectory + @"Resources\";
        public const float SlideSize16x9Width = 33.867f;
        public const float SlideSizeHeight = 19.05f;
        public const float SlideSize4x3Width = 25.4f;
        public const float CoverImageWidth = 27.6f;
        public const float CoverImageHeight = 9.4f;
        public const float CoverImageTop = 4.0f;
        public const float CoverImageLeft = 3.13f;
        public const float CoverLogoWidth = 3.84f;
        public const float CoverLogoHeight = 3.37f;
        public const float CoverLogoTop = 2.71f;
        public const float CoverLogoLeft = 15.01f;
        public const float CoverCommentWidth = 16.86f;
        public const float CoverCommentHeight = 4.02f;
        public const float CoverCommentTop = 12.53f;
        public const float CoverCommentLeft = 8.5f;
        public const float PaneHeight = 14.32f;
        public const float HorizontalMargin = 0.26f;
        public const float VerticalMargin = 0.13f;

        public const string FontNanumSquareBold = "나눔스퀘어 Bold";
        public const string FontNanumSquareExBold = "나눔스퀘어라운드 ExtraBold";
        public const string FontMalgunGothic = "맑은 고딕";
        public const string FontKopubDotumBold = "Kopub돋움체 Bold";
        public const string FontSequenceTitle = "KT&G 상상제목 B";
        public const string FontSongTitle = "HY궁서B";
        public const float RoundedRectangleRadius = 0.07707f;
    }

    public enum SlideSizeType { WideScreen, Normal }
    public enum LyricsVAlignmentType { Top, Center, Bottom }
    public enum TextEmphasis
    {
        None = 0b_0000_0000,
        Bold = 0b_0000_0001,
        Italic = 0b_0000_0010,
        UnderLine = 0b_0000_0100,
        Shadow = 0b_0000_1000,
        StrikeThrough = 0b_0001_0000
    }

    public class ShadowOptions
    {
        public ShadowOptions(MsoShadowStyle style = MsoShadowStyle.msoShadowStyleOuterShadow,
            int color = 0, float transparency = 0.57f, float size = 100, float blur = 3, double angle = 45, float distance = 3)
        {
            Style = style;
            Color = color;
            Transparency = transparency;
            Size = size;
            Blur = blur;
            Angle = angle;
            Distance = distance;
        }
        public MsoShadowStyle Style { get; set; }
        public int Color { get; set; }
        public float Transparency { get; set; }
        public float Size { get; set; }
        public float Blur { get; set; }
        public double Angle { get; set; }
        public float Distance { get; set; }
    }
}
