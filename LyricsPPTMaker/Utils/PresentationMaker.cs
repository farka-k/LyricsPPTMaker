using LyricsPPTMaker.Models;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System;

namespace LyricsPPTMaker
{
    public static class PresentationMaker
    {
        public static List<string> SplitLyrics(string lyrics){
            if (!lyrics.Contains('\r'))
            {
                lyrics = lyrics.Replace("\n", "\r\n");
            }
            lyrics = lyrics.Replace("\r\n\r\n", "\r\n \r\n");
            List<string> phrases = lyrics.Split("\r\n \r\n").ToList();
            return phrases;
        }

        public static void MakeLyricsSlides(List<string> phrases, LyricSlideOptions options, bool isCopy)
        {
            PowerPoint.Application pptApp = new PowerPoint.Application();
            PowerPoint.Presentations pptPres = pptApp.Presentations;
            PowerPoint.Presentation presentation = pptPres.Add(MsoTriState.msoFalse);
            if (options.SizeType == SlideSizeType.Normal)
                presentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeOnScreen;
            else
            {
                presentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeCustom;
                presentation.PageSetup.SlideWidth = Utils.CMToPoint(options.SlideWidth);
                presentation.PageSetup.SlideHeight = Utils.CMToPoint(options.SlideHeight);
            }

            presentation.SlideMaster.Background.Fill.ForeColor.RGB = options.BackgroundColor;

            PowerPoint.CustomLayout pcl = presentation.SlideMaster.CustomLayouts[7];
            presentation.Slides.AddSlide(1, pcl);

            //first slide
            int lastSlide = 1;
            presentation.Slides[lastSlide].Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                0, Utils.CMToPoint(options.Offset),
                Utils.CMToPoint(options.SlideWidth), 0);
            var currentShape = presentation.Slides[lastSlide].Shapes[1];
            SetTextEffectOptions(ref currentShape, phrases[0], options.FontName, options.FontSize,
                options.Emphasis, options.FontColor, MsoParagraphAlignment.msoAlignCenter, options.VerticalAlignment, lineSpace: 1.5f);

            float offset = Utils.CMToPoint(options.Offset);
            if (options.VerticalAlignment== MsoVerticalAnchor.msoAnchorMiddle)
            {
                offset = offset + (Utils.CMToPoint(Constants.SlideSizeHeight) - currentShape.Height) / 2;
            }
            else if (options.VerticalAlignment==MsoVerticalAnchor.msoAnchorBottom)
            {
                offset = offset + Utils.CMToPoint(Constants.SlideSizeHeight) - currentShape.Height;
            }
            currentShape.Top = offset;

            //remain slides
            for (int i = phrases.Count - 1; i > 0; --i)
            {
                presentation.Slides[lastSlide].Duplicate();
            }
            for (int i = 1; i < phrases.Count; ++i)
            {
                presentation.Slides[i + 1].Shapes[1].TextFrame2.TextRange.Text = phrases[i].Trim();
            }

            //copy all
            presentation.Slides.Range().Copy();

            if (isCopy)
                MessageBox.Show("슬라이드가 클립보드에 복사되었습니다.");
            presentation.Close();
        }

        public static void SetTextEffectOptions(ref PowerPoint.Shape currentShape, string text, string fontName = "굴림", float fontSize = 18,
            TextEmphasis emphasis = TextEmphasis.None, int fontFillColor = 0,
            MsoParagraphAlignment paragraphAlignment = MsoParagraphAlignment.msoAlignLeft,
            MsoVerticalAnchor verticalAnchor = MsoVerticalAnchor.msoAnchorTop,
            MsoAutoSize autoSize = MsoAutoSize.msoAutoSizeShapeToFitText,
            float lineSpace = 1, MsoTriState fontLineVisible = MsoTriState.msoFalse,
            int fontLineColor = 0, float fontLineWeight = 1, ShadowOptions shadowOptions = null)
        {
            currentShape.TextFrame2.AutoSize = autoSize;
            currentShape.TextFrame2.TextRange.Text = text;
            currentShape.TextFrame2.TextRange.Font.NameFarEast = fontName;
            currentShape.TextFrame2.TextRange.Font.NameAscii = "(한글 글꼴 사용)";
            currentShape.TextFrame2.TextRange.Font.Size = fontSize;
            if ((emphasis & TextEmphasis.Bold) == TextEmphasis.Bold) currentShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
            else currentShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoFalse;
            if ((emphasis & TextEmphasis.Italic) == TextEmphasis.Italic) currentShape.TextFrame2.TextRange.Font.Italic = MsoTriState.msoTrue;
            else currentShape.TextFrame2.TextRange.Font.Italic = MsoTriState.msoFalse;
            if ((emphasis & TextEmphasis.UnderLine) == TextEmphasis.UnderLine) currentShape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
            else currentShape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
            if ((emphasis & TextEmphasis.StrikeThrough) == TextEmphasis.StrikeThrough) currentShape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
            else currentShape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoFalse;
            if ((emphasis & TextEmphasis.Shadow) == TextEmphasis.Shadow)
            {
                var textFrame = currentShape.TextFrame2;
                if (shadowOptions == null) shadowOptions = new ShadowOptions();
                AdjustShadow(ref textFrame, shadowOptions.Color, shadowOptions.Transparency, shadowOptions.Size,
                    shadowOptions.Blur, shadowOptions.Angle, shadowOptions.Distance, shadowOptions.Style);
            }

            currentShape.TextFrame2.TextRange.ParagraphFormat.Alignment = paragraphAlignment;
            currentShape.TextFrame2.VerticalAnchor = verticalAnchor;
            currentShape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = lineSpace;
            currentShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontFillColor;
            currentShape.TextFrame2.TextRange.Font.Line.Visible = fontLineVisible;
            if (fontLineVisible == MsoTriState.msoTrue)
            {
                currentShape.TextFrame2.TextRange.Font.Line.ForeColor.RGB = fontLineColor;
                currentShape.TextFrame2.TextRange.Font.Line.Weight = fontLineWeight;
            }
        }

        public static void AdjustShadow(ref PowerPoint.TextFrame2 textFrame,
            int color = 0, float transparency = 0.57f, float size = 100, float blur = 3, double angle = 45, float distance = 3,
            MsoShadowStyle shadowStyle = MsoShadowStyle.msoShadowStyleOuterShadow)
        {
            textFrame.TextRange.Font.Shadow.Style = shadowStyle;
            textFrame.TextRange.Font.Shadow.ForeColor.RGB = color;
            textFrame.TextRange.Font.Shadow.Transparency = transparency;
            textFrame.TextRange.Font.Shadow.Size = size;
            textFrame.TextRange.Font.Shadow.Blur = blur;
            textFrame.TextRange.Font.Shadow.OffsetX = (float)(distance * Math.Cos(angle));
            textFrame.TextRange.Font.Shadow.OffsetY = (float)(distance * Math.Sin(angle));
        }

        public static void AdjustShadow(ref PowerPoint.Shape shape,
            int color = 0, float transparency = 0.6f, float size = 100, float blur = 4, double angle = 45, float distance = 3,
            MsoShadowStyle shadowStyle = MsoShadowStyle.msoShadowStyleOuterShadow)
        {
            shape.Shadow.Style = shadowStyle;
            shape.Shadow.ForeColor.RGB = color;
            shape.Shadow.Transparency = transparency;
            shape.Shadow.Size = size;
            shape.Shadow.Blur = blur;
            shape.Shadow.OffsetX = (float)(distance * Math.Cos(angle));
            shape.Shadow.OffsetY = (float)(distance * Math.Sin(angle));
        }
    }
}
