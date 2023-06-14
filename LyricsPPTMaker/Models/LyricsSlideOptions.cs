using Microsoft.Office.Core;

namespace LyricsPPTMaker.Models
{
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
    }
}
