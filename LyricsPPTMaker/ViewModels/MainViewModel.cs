using CommunityToolkit.Mvvm.ComponentModel;
using LyricsPPTMaker.Bases;
using System.Collections.Generic;
using DT = System.Drawing.Text;
using CommunityToolkit.Mvvm.Input;
using System.Windows;
using System;
using System.Text.Json;
using LyricsPPTMaker.Models;
using System.IO;
using System.Net;
using Microsoft.Office.Core;
using System.Globalization;
using System.Collections.ObjectModel;
using System.Linq;

namespace LyricsPPTMaker.ViewModels
{
    public partial class MainViewModel : ViewModelBase
    {
        public MainViewModel()
        {
            Title = "LyricsPPT Maker";

            searchboxText = "제목을 입력하세요";
            usagePopupOpen = false;
            songList = new ObservableCollection<SongInfo>();
            currentSongListIndex = -1;
            currentSelectedSong = null;
            totalSongs = 0;
            backgroundColor = "#FF100010";
            foregroundColor = "#FFFFFFFF";
            fontSelectedIndex = 0;

            sizeType = SlideSizeType.WideScreen;
            slideWidth = Constants.SlideSize16x9Width;
            slideHeight = Constants.SlideSizeHeight;
            GetFontFamilies();
            fontSize = 40;
            isBold = false;
            isItalic = false;
            isUnderline = false;
            fontPreviewText = "AaBbCc가나다라1234";

            lyricsVAlignment = LyricsVAlignmentType.Top;
            alignmentOffset = 0.00f;
            offsetMin = 0;
            offsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                - (Utils.PointToCM(fontSize) + Constants.VerticalMargin * 2 + 0.56));

            includeTitleMark = false;
            previewTextPadding = new Thickness(0, Utils.CMToPixel(alignmentOffset.Value + 0.7) / 4, 0, 0);
        }

        [ObservableProperty]
        private string? searchboxText;
        [ObservableProperty]
        private ObservableCollection<SongInfo> songList;
        [ObservableProperty]
        private int currentSongListIndex;
        //[ObservableProperty]
        //private bool? listBoxIsSelected;
        [ObservableProperty]
        private object? currentSelectedSong;

        [ObservableProperty]
        private int totalSongs;
        [ObservableProperty]
        private string? lyricsboxText;
        [ObservableProperty]
        private bool usagePopupOpen;

        [ObservableProperty]
        private SlideSizeType sizeType;
        [ObservableProperty]
        private float slideWidth;
        [ObservableProperty]
        private float slideHeight;

        [ObservableProperty]
        private string backgroundColor;
        [ObservableProperty]
        private string foregroundColor;

        [ObservableProperty]
        private List<string>? fontList;
        [ObservableProperty]
        private int fontSelectedIndex;
        [ObservableProperty]
        private float fontSize;
        [ObservableProperty]
        private bool isBold;
        [ObservableProperty]
        private bool isItalic;
        [ObservableProperty]
        private bool isUnderline;

        [ObservableProperty]
        private string fontPreviewText;
        [ObservableProperty]
        private LyricsVAlignmentType lyricsVAlignment;
        [ObservableProperty]
        private float? alignmentOffset;
        [ObservableProperty]
        private decimal? offsetMin;
        [ObservableProperty]
        private decimal? offsetMax;
        [ObservableProperty]
        private bool includeTitleMark;

        [ObservableProperty]
        private bool previewPopupOpen = false;
        [ObservableProperty]
        private Thickness previewTextPadding;

        private void GetFontFamilies()
        {
            FontList = new List<string>();
            using (DT.InstalledFontCollection col = new DT.InstalledFontCollection())
            {
                foreach (var font in col.Families)
                {
                    FontList.Add(font.Name);
                }
            }
            FontList.Sort();
        }

        [RelayCommand]
        public void SearchLyrics()
        {
            if (string.IsNullOrWhiteSpace(SearchboxText) ||
                SearchboxText == "제목을 입력하세요")
            {
                MessageBox.Show("제목을 입력하세요");
                return;
            }
            //Clear SongInfo List
            SongList.Clear();
            TotalSongs = 0;
            CurrentSongListIndex = -1;

            //검색
            //google custom search api
            string response = string.Empty;
            string? key = Environment.GetEnvironmentVariable("GoogleCustomSearchKey", EnvironmentVariableTarget.User);
            string? cx = Environment.GetEnvironmentVariable("GoogleCustomSearchEngineID", EnvironmentVariableTarget.User);
            int maxSearchCount = 5;
            string requestUrl = "https://www.googleapis.com/customsearch/v1?" + "key=" + key + "&cx=" + cx +
                "&q=" + SearchboxText + " 가사" + "&num=" + maxSearchCount.ToString();

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl);
            request.Method = "GET";
            request.ContentType = "application/json";
            using (HttpWebResponse resp = (HttpWebResponse)request.GetResponse())
            {
                HttpStatusCode status = resp.StatusCode;
                if (status == HttpStatusCode.OK)
                {
                    Stream respStream = resp.GetResponseStream();
                    using (StreamReader sr = new StreamReader(respStream))
                    {
                        response = sr.ReadToEnd();
                    }
                }
            }
            var responseObj = JsonSerializer.Deserialize<GoogleCustomSearchAPIResponseModel>(response);
            TotalSongs = responseObj.items.Count;
            for (int i = 0; i < TotalSongs; ++i)
            {
                string? lyricsUrl = responseObj.items[i].link;

                var songInfo = LyricsCollector.GetLyrics(lyricsUrl);
                SongList.Add(songInfo);
            }

            if (TotalSongs > 0) CurrentSongListIndex = 0;
            PreviewPopupOpen = true;
        }

        [RelayCommand]
        public void SearchboxGotFocus()
        {
            if (SearchboxText == "제목을 입력하세요")
            {
                SearchboxText = "";
            }
        }

        [RelayCommand]
        public void SearchboxLostFocus()
        {
            if (SearchboxText == "")
            {
                SearchboxText = "제목을 입력하세요";
            }
        }

        [RelayCommand]
        public void UsageMouseEnter()
        {
            UsagePopupOpen = true;
        }

        [RelayCommand]
        public void UsageMouseLeave()
        {
            UsagePopupOpen = false;
        }

        [RelayCommand]
        public void ChangePreviewStatus()
        {
            PreviewPopupOpen = !PreviewPopupOpen;
        }

        [RelayCommand]
        public void GetPreviousSongInfo()
        {
            CurrentSongListIndex = (CurrentSongListIndex - 1 >= 0) ? CurrentSongListIndex - 1 : SongList.Count - 1;
        }

        [RelayCommand]
        public void GetNextSongInfo()
        {
            CurrentSongListIndex = (CurrentSongListIndex + 1 < SongList.Count) ? CurrentSongListIndex + 1 : 0;
        }

        [RelayCommand]
        public void VAlignmentChanged()
        {
            if (LyricsVAlignment == LyricsVAlignmentType.Top)
            {
                OffsetMin = 0;
                OffsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                    - (Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56));
            }
            else if (LyricsVAlignment == LyricsVAlignmentType.Center)
            {
                OffsetMin = (decimal?)(-Constants.SlideSizeHeight / 2
                + Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56);
                OffsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                    - (Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56));
            }
            else    //LyricsVAlignment == LyricsVAlignmentType.Bottom
            {
                OffsetMin = (decimal?)(-Constants.SlideSizeHeight / 2
                + Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56);
                OffsetMax = 0;
            }
        }

        [RelayCommand]
        public void SongListIndexChanged()
        {
            if (CurrentSongListIndex < 0) return;
            CurrentSelectedSong = SongList[CurrentSongListIndex];
            LyricsboxText = ((SongInfo)CurrentSelectedSong).Lyrics;
        }

        [RelayCommand]
        public void CopySlide()
        {   
            PreviewPopupOpen = false;
            var phrases = PresentationMaker.SplitLyrics(LyricsboxText);
            TextEmphasis emphasis = TextEmphasis.None;
            if (IsBold == true)
            {
                emphasis |= TextEmphasis.Bold;
            }
            if (IsItalic == true)
            {
                emphasis |= TextEmphasis.Italic;
            }
            if (IsUnderline == true)
            {
                emphasis |= TextEmphasis.UnderLine;
            }

            MsoVerticalAnchor vAlignment;
            if (LyricsVAlignment == LyricsVAlignmentType.Top)
                vAlignment = MsoVerticalAnchor.msoAnchorTop;
            else if (LyricsVAlignment == LyricsVAlignmentType.Center)
                vAlignment = MsoVerticalAnchor.msoAnchorMiddle;
            else
                vAlignment = MsoVerticalAnchor.msoAnchorBottom;


            var options = new LyricSlideOptions
            {
                SizeType = SizeType,
                SlideWidth = SlideWidth,
                SlideHeight = SlideHeight,
                BackgroundColor = int.Parse(BackgroundColor.Remove(0, 3), NumberStyles.HexNumber),
                FontColor = int.Parse(ForegroundColor.Remove(0, 3), NumberStyles.HexNumber),
                Emphasis = emphasis,
                FontName = FontList[FontSelectedIndex],
                FontSize = FontSize,
                Offset = AlignmentOffset.Value,
                VerticalAlignment = vAlignment
            };
            PresentationMaker.MakeLyricsSlides(phrases, options, true);
        }

        [RelayCommand]
        public void Reset()
        {
            SearchboxText = "제목을 입력하세요";
            UsagePopupOpen = false;
            SongList.Clear();
            CurrentSongListIndex = -1;
            TotalSongs = 0;

            SizeType = SlideSizeType.WideScreen;
            SlideWidth = Constants.SlideSize16x9Width;
            BackgroundColor = "#FF100010";
            ForegroundColor = "#FFFFFFFF";
            FontSelectedIndex = 0;
            FontSize = 40;
            IsBold = false;
            IsItalic = false;
            IsUnderline = false;
            FontPreviewText = "AaBbCc가나다라1234";

            LyricsVAlignment = LyricsVAlignmentType.Top;
            AlignmentOffset = 0;
            OffsetMin = 0;
            OffsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                - (Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56));
            IncludeTitleMark = false;
            PreviewPopupOpen = false;
            PreviewTextPadding = new Thickness(0, Utils.CMToPixel(AlignmentOffset.Value + 0.7) / 4, 0, 0);
        }
    }
}
