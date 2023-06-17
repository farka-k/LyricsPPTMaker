using Microsoft.Office.Core;
using System;

namespace LyricsPPTMaker.Models
{
    [Serializable]
    public class LyricSlideOptions
    {
        public SlideSizeType SizeType { get; set; }
        public float SlideWidth { get; set; }
        public float SlideHeight { get; set; }
        public int BackgroundColor { get; set; }
        public int FontColor { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public TextEmphasis Emphasis { get; set; }
        public MsoVerticalAnchor VerticalAlignment { get; set; }
        public float Offset { get; set; }

        public LyricSlideOptions() {
            SizeType = SlideSizeType.WideScreen;
            SlideWidth = Constants.SlideSize16x9Width;
            SlideHeight = Constants.SlideSizeHeight;
            BackgroundColor = 1_048_592;
            FontColor = 16_777_215;
            FontName = "굴림";
            FontSize = 40;
            Emphasis = TextEmphasis.None;
            VerticalAlignment = MsoVerticalAnchor.msoAnchorTop;
            Offset = 0;
        }
        public LyricSlideOptions(SlideSizeType sizetype=SlideSizeType.WideScreen, 
            float width=Constants.SlideSize16x9Width, float height=Constants.SlideSizeHeight,
            int background=1_048_592, int foreground=16_777_215, string fontname="굴림", float fontsize=40,
            TextEmphasis emphasis=TextEmphasis.None, MsoVerticalAnchor valign=MsoVerticalAnchor.msoAnchorTop, float offset=0)
        {
            SizeType = sizetype;
            SlideWidth = width;
            SlideHeight = height;
            BackgroundColor = background;
            FontColor = foreground;
            FontName = fontname;
            FontSize = fontsize;
            Emphasis = emphasis;
            VerticalAlignment = valign;
            Offset = offset;
        }

        public LyricSlideOptions(LyricSlideOptions options)
        {
            SizeType=options.SizeType;
            SlideWidth = options.SlideWidth;
            SlideHeight = options.SlideHeight;
            BackgroundColor = options.BackgroundColor;
            FontColor = options.FontColor;
            FontName = options.FontName;
            FontSize = options.FontSize;
            Emphasis = options.Emphasis;
            VerticalAlignment = options.VerticalAlignment;
            Offset = options.Offset;
        }

        public void CopyProperty(LyricSlideOptions options)
        {
            SizeType = options.SizeType;
            SlideWidth = options.SlideWidth;
            SlideHeight = options.SlideHeight;
            BackgroundColor = options.BackgroundColor;
            FontColor = options.FontColor;
            FontName = options.FontName;
            FontSize = options.FontSize;
            Emphasis = options.Emphasis;
            VerticalAlignment = options.VerticalAlignment;
            Offset = options.Offset;
        }
    } 
}
