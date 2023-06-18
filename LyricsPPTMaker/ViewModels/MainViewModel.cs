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
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.Win32;

namespace LyricsPPTMaker.ViewModels
{
    public partial class MainViewModel : ViewModelBase
    {
        public MainViewModel()
        {
            Title = "LyricsPPT Maker";
            usagePopupOpen = false;
            songList = new ObservableCollection<SongInfo>();
            currentSongListIndex = -1;
            currentSelectedSong = null;
            totalSongs = 0;
            GetFontFamilies();
            presetList = new ObservableCollection<Preset>();
            InitPreset();
            currentSelectedPresetIndex = -1;
            newPresetDialogOpen = false;
            presetNameDialogOpen = false;
            presetManagerOpen = false;
            presetRenameDialogOpen = false;
            newPresetDialogResult = PresetDialogResult.None;
            presetGenMethod = PresetGenerationMethod.None;
            selectedPresetManagerIndex = -1;
            newPresetName = String.Empty;
            isDuplicated = false;
            isEmptyOrWhiteSpace = false;
            fontPreviewText = "AaBbCc가나다라1234";
            previewPopupOpen = false;
            InitInputs();
            isSearchboxTextFocus = true;
        }

        #region ViewModel Property
        #region LyricsSearch
        [ObservableProperty]
        private string? searchboxText;
        [ObservableProperty]
        private bool isSearchboxTextFocus;
        [ObservableProperty]
        private ObservableCollection<SongInfo> songList;
        [ObservableProperty]
        private int currentSongListIndex;
        [ObservableProperty]
        private object? currentSelectedSong;

        [ObservableProperty]
        private int totalSongs;
        [ObservableProperty]
        private string? lyricsboxText;
        [ObservableProperty]
        private bool usagePopupOpen;
        #endregion LyricsSearch

        #region Preset
        private PresetXMLSerializer PresetSerializer;
        [ObservableProperty]
        private ObservableCollection<Preset> presetList;
        [ObservableProperty]
        private int currentSelectedPresetIndex;
        
        [ObservableProperty]
        private bool newPresetDialogOpen;
        [ObservableProperty]
        private bool presetNameDialogOpen;
        [ObservableProperty]
        private bool presetManagerOpen;
        [ObservableProperty]
        private bool presetRenameDialogOpen;
        [ObservableProperty]
        private PresetDialogResult newPresetDialogResult;
        [ObservableProperty]
        private PresetGenerationMethod presetGenMethod;
        [ObservableProperty]
        private string newPresetName;
        [ObservableProperty]
        private bool isNewPresetNameFocus;

        [ObservableProperty]
        private int selectedPresetManagerIndex;
        [ObservableProperty]
        private bool isDuplicated;
        [ObservableProperty]
        private bool isEmptyOrWhiteSpace;
        #endregion Preset

        #region SlideOptions
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
        #endregion SlideOptions

        #region Preview
        [ObservableProperty]
        private bool previewPopupOpen = false;
        [ObservableProperty]
        private Thickness previewTextPadding;
        #endregion Preview

        #endregion ViewModel Property

        /// <summary>
        /// Load System Installed Fonts.
        /// </summary>
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

        /// <summary>
        /// Set initial values to user input components.
        /// </summary>
        private void InitInputs()
        {
            SizeType = SlideSizeType.WideScreen;
            SlideWidth = Constants.SlideSize16x9Width;
            SlideHeight = Constants.SlideSizeHeight;
            BackgroundColor = "#FF100010";
            ForegroundColor = "#FFFFFFFF";
            FontSelectedIndex = 0;
            FontSize = 40;
            IsBold = false;
            IsItalic = false;
            IsUnderline = false;

            LyricsVAlignment = LyricsVAlignmentType.Top;
            AlignmentOffset = 0.00f;
            OffsetMin = 0;
            OffsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                - (Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56));

            PreviewTextPadding = new Thickness(0, Utils.CMToPixel(AlignmentOffset.Value + 0.7) / 4, 0, 0);
        }

        /// <summary>
        /// Load Preset from Xml File. Generate new one if not exists.
        /// </summary>
        private void InitPreset()
        {
            PresetSerializer = new PresetXMLSerializer();
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + @"\Presets.xml"))
            {
                PresetSerializer.SetPresetList(PresetList);
                PresetSerializer.Serialize();
                return;
            }
            PresetList=PresetSerializer.Deserialize(AppDomain.CurrentDomain.BaseDirectory+@"Presets.xml");
        }

        private void UpdatePresetFile()
        {
            PresetSerializer.Serialize();
        }

        /// <summary>
        /// Search lyrics.
        /// </summary>
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
            string? key, cx; int maxSearchCount;
            using (RegistryKey SearchKey = Registry.CurrentUser.OpenSubKey(@"Software\Lazyworks\LyricsPPTMaker"))
            {
                if (SearchKey == null)
                {
                    MessageBox.Show("Can't find Search Api keys.\nPlease Reinstall the application.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                key = SearchKey.GetValue("GoogleCustomSearchKey") as string;
                cx = SearchKey.GetValue("GoogleCustomSearchEngineID") as string;
                if (String.IsNullOrWhiteSpace(key) || String.IsNullOrWhiteSpace(cx)) return;
                Object? SearchCountKey = SearchKey.GetValue("SearchCount");
                maxSearchCount = (SearchCountKey != null) ? (int)SearchCountKey : 5;
            }
            StringBuilder requestUrl = new StringBuilder();
            requestUrl.Append("https://www.googleapis.com/customsearch/v1?key=");
            requestUrl.Append(key);
            requestUrl.Append("&cx=");
            requestUrl.Append(cx);
            requestUrl.Append("&q=");
            requestUrl.Append(SearchboxText);
            requestUrl.Append(" 가사&num=");
            requestUrl.Append(maxSearchCount.ToString());

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestUrl.ToString());
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

        /// <summary>
        /// Make disappear of guide message in search box.
        /// </summary>
        [RelayCommand]
        public void SearchboxGotFocus()
        {
            if (SearchboxText == "제목을 입력하세요")
            {
                SearchboxText = "";
            }
        }

        /// <summary>
        /// Show guide message in search box when string is empty or white spaces.
        /// </summary>
        [RelayCommand]
        public void SearchboxLostFocus()
        {
            if (String.IsNullOrWhiteSpace(SearchboxText))
            {
                SearchboxText = "제목을 입력하세요";
            }
        }

        /// <summary>
        /// Show Usage Popup if mouse is over the button.
        /// </summary>
        [RelayCommand]
        public void UsageMouseEnter()
        {
            UsagePopupOpen = true;
        }

        /// <summary>
        /// Hide Usage Popup if mouse leave the button.
        /// </summary>
        [RelayCommand]
        public void UsageMouseLeave()
        {
            UsagePopupOpen = false;
        }

        /// <summary>
        /// Get previous song.
        /// </summary>
        [RelayCommand]
        public void GetPreviousSongInfo()
        {
            CurrentSongListIndex = (CurrentSongListIndex - 1 >= 0) ? CurrentSongListIndex - 1 : SongList.Count - 1;
        }

        /// <summary>
        /// Get next song.
        /// </summary>
        [RelayCommand]
        public void GetNextSongInfo()
        {
            CurrentSongListIndex = (CurrentSongListIndex + 1 < SongList.Count) ? CurrentSongListIndex + 1 : 0;
        }

        /// <summary>
        /// Change Offset range when VerticalAlignment is changed.
        /// </summary>
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

        /// <summary>
        /// Set current song's lyric to lyric textbox.
        /// </summary>
        [RelayCommand]
        public void SongListIndexChanged()
        {
            if (CurrentSongListIndex < 0) return;
            CurrentSelectedSong = SongList[CurrentSongListIndex];
            LyricsboxText = ((SongInfo)CurrentSelectedSong).Lyrics;
        }

        [RelayCommand]
        public void OpenNewPresetDialog()
        {
            InitDialog();
        }

        /// <summary>
        /// Close NewPresetDialog without any change.
        /// </summary>
        [RelayCommand]
        public void CloseNewPresetDialog()
        {
            NewPresetDialogResult = PresetDialogResult.Cancel;
            NewPresetDialogOpen = false;
        }

        /// <summary>
        /// Close PresetNameDialog without any change.
        /// </summary>
        [RelayCommand]
        public void ClosePresetNameDialog()
        {
            NewPresetName = String.Empty;
            PresetNameDialogOpen = false;
            IsNewPresetNameFocus = false;
        }

        /// <summary>
        /// Show Preset Manager Window
        /// </summary>
        [RelayCommand]
        public void OpenPresetManager()
        {
            PresetManagerOpen = true;
        }

        /// <summary>
        /// Close Preset Manager Window.
        /// </summary>
        [RelayCommand]
        public void ClosePresetManager()
        {
            PresetManagerOpen = false;
        }

        /// <summary>
        /// Add Preset with default name and options. This function only execute in the PresetManager.
        /// </summary>
        [RelayCommand]
        public void AddDefaultPreset()
        {
            PresetList.Add(new Preset(GenerateDefaultPresetName(), new LyricSlideOptions()));
            UpdatePresetFile();
        }

        /// <summary>
        /// Generate default preset name.
        /// </summary>
        private string GenerateDefaultPresetName()
        {
            List<string> presetNameMatchedList = new List<string>();
            Regex rxName = new Regex(@"^Preset [\d]+$");
            Regex rxNumber = new Regex(@"[\d]+");
            string name;
            foreach (var preset in PresetList)
            {
                if (rxName.IsMatch(preset.Name))
                    presetNameMatchedList.Add(preset.Name);
            }
            if (presetNameMatchedList.Count > 0)
            {
                presetNameMatchedList.Sort();
                int number = int.Parse(rxNumber.Match(presetNameMatchedList.Last()).Value) + 1;
                name = rxNumber.Replace(presetNameMatchedList.Last(), number.ToString());
            }
            else name = "Preset 1";
            return name;
        }

        /// <summary>
        /// Copy selected Preset. This function only execute in the PresetManager.
        /// </summary>
        [RelayCommand]
        public void CopySelectedPreset()
        {
            if (SelectedPresetManagerIndex < 0) return;
            Preset original = PresetList[SelectedPresetManagerIndex];
            string name = original.Name + "_Copy";
            PresetList.Insert(SelectedPresetManagerIndex + 1, 
                new Preset(name, new LyricSlideOptions(original.Options)));
            UpdatePresetFile();
        }


        [RelayCommand]
        public void RenameSelectedPreset()
        {
            if (SelectedPresetManagerIndex < 0) return;
            NewPresetName = PresetList[SelectedPresetManagerIndex].Name;
            IsDuplicated = false;
            IsEmptyOrWhiteSpace = false;
            PresetRenameDialogOpen = true;
            IsNewPresetNameFocus = true;
        }

        [RelayCommand]
        public void CloseRenameDialog()
        {
            PresetRenameDialogOpen = false;
            IsNewPresetNameFocus = false;
        }

        [RelayCommand]
        public void UpdatePresetName()
        {
            if (IsDuplicated || IsEmptyOrWhiteSpace) return;
            if (PresetList[SelectedPresetManagerIndex].Name == NewPresetName)
            {
                CloseRenameDialog();
                return;
            }
            int index = SelectedPresetManagerIndex;
            var options = new LyricSlideOptions(PresetList[index].Options);
            PresetList.Insert(index+1, new Preset(NewPresetName, options));
            PresetList.RemoveAt(index);
            SelectedPresetManagerIndex = index;
            PresetRenameDialogOpen = false;
            IsNewPresetNameFocus = false;
            UpdatePresetFile();
        }

        [RelayCommand]
        public void ValidatePresetName(object param)
        {
            string name = (string)param;
            IsDuplicated = IsDuplicatedName(name, SelectedPresetManagerIndex);
            IsEmptyOrWhiteSpace = String.IsNullOrWhiteSpace(name);
            return;
        }

        /// <summary>
        /// Delete selected Preset. This function only execute in the PresetManager.
        /// </summary>
        [RelayCommand]
        public void DeleteSelectedPreset()
        {
            if (SelectedPresetManagerIndex < 0) return;
            int index = SelectedPresetManagerIndex;
            PresetList.RemoveAt(index);
            if (index >= PresetList.Count)  //delete last item
                SelectedPresetManagerIndex = index - 1;
            else SelectedPresetManagerIndex = index;
            UpdatePresetFile();
        }

        [RelayCommand]
        public void ItemOrderUp()
        {
            MoveItem(PresetList, -1);

        }

        [RelayCommand]
        public void ItemOrderDown()
        {
            MoveItem(PresetList, 1);
        }

        private void MoveItem(ObservableCollection<Preset> list,int direction)
        {
            if (SelectedPresetManagerIndex < 0) 
                return;     //No Selected Item
            int newIdx = SelectedPresetManagerIndex + direction;
            if (newIdx < 0 || newIdx>=list.Count)
                return;     //Index Out of Range
            var selectedItem = list[SelectedPresetManagerIndex];
            list.Remove(selectedItem);
            list.Insert(newIdx, selectedItem);
            SelectedPresetManagerIndex = newIdx;
            UpdatePresetFile();
        }

        private void InitDialog()
        {
            NewPresetDialogResult = PresetDialogResult.None;
            PresetGenMethod = PresetGenerationMethod.None;
            NewPresetName = string.Empty;
            NewPresetDialogOpen = true;
            
        }

        [RelayCommand]
        public void SetGenerationMethod(string param)
        {
            if (param == "Copy")
            {
                PresetGenMethod = PresetGenerationMethod.CopyFromCurrent;

            }
            else //param=="Default"
            {
                PresetGenMethod = PresetGenerationMethod.Default;

            }
            NewPresetDialogResult = PresetDialogResult.Selected;
            NewPresetDialogOpen = false;
            PresetNameDialogOpen = true;
            IsNewPresetNameFocus = true;
        }

        private LyricSlideOptions GetSlideOptionsFromCurrentSetting
            () => new LyricSlideOptions()
            {
                SizeType = SizeType,
                SlideWidth = SlideWidth,
                SlideHeight = SlideHeight,
                BackgroundColor = GetColorInteger(BackgroundColor),
                FontColor = GetColorInteger(ForegroundColor),
                FontName = FontList[FontSelectedIndex],
                FontSize = FontSize,
                Emphasis = GetTextEmphasis(IsBold, IsItalic, IsUnderline),
                VerticalAlignment = GetMsVAnchor(LyricsVAlignment),
                Offset = AlignmentOffset.Value
            };

        [RelayCommand]
        public void AddNewPreset()
        {
            if (IsDuplicatedName(NewPresetName, CurrentSelectedPresetIndex))
            {
                MessageBox.Show("중복된 이름입니다.", "Error: DuplicatedName", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (String.IsNullOrWhiteSpace(NewPresetName))
            {
                MessageBox.Show("빈 문자열이나 공백으로 이름지을 수 없습니다.", "Error: EmptyString", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Preset newPreset;
            if (PresetGenMethod == PresetGenerationMethod.CopyFromCurrent)
            {
                newPreset = new Preset(
                    NewPresetName,
                    GetSlideOptionsFromCurrentSetting()
                );
            }
            else
            {
                newPreset = new Preset(
                    NewPresetName,
                    new LyricSlideOptions()
                );
            }
            PresetList.Add(newPreset);
            CurrentSelectedPresetIndex = PresetList.Count - 1;
            PresetNameDialogOpen = false;
            IsNewPresetNameFocus = false;
            UpdatePresetFile();
        }

        [RelayCommand]
        public void UpdateInputs()
        {
            SelectedPresetManagerIndex = CurrentSelectedPresetIndex;

            if (CurrentSelectedPresetIndex == -1)
            {
                InitInputs();
            }
            else
            {
                var currentPresetItem = PresetList[CurrentSelectedPresetIndex];
                var currentPreset = currentPresetItem.Options;

                SizeType = currentPreset.SizeType;
                SlideWidth = currentPreset.SlideWidth;
                SlideHeight = currentPreset.SlideHeight;
                BackgroundColor = GetColorString(currentPreset.BackgroundColor);
                ForegroundColor = GetColorString(currentPreset.FontColor);
                FontSelectedIndex = FontList.IndexOf(currentPreset.FontName);
                FontSize = currentPreset.FontSize;
                SetEmphasis(currentPreset.Emphasis);
                LyricsVAlignment = GetLyricVAlignment(currentPreset.VerticalAlignment);
                AlignmentOffset = currentPreset.Offset;
                /*OffsetMin = 0;
                OffsetMax = (decimal?)(Constants.SlideSizeHeight / 2
                    - (Utils.PointToCM(FontSize) + Constants.VerticalMargin * 2 + 0.56));*/

                PreviewTextPadding = new Thickness(0, Utils.CMToPixel(AlignmentOffset.Value + 0.7) / 4, 0, 0);

            }
        }

        private MsoVerticalAnchor GetMsVAnchor(LyricsVAlignmentType lVAlign)
        {
            MsoVerticalAnchor MsVAnchor;
            if (lVAlign == LyricsVAlignmentType.Top)
                MsVAnchor = MsoVerticalAnchor.msoAnchorTop;
            else if (lVAlign == LyricsVAlignmentType.Center)
                MsVAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            else
                MsVAnchor = MsoVerticalAnchor.msoAnchorBottom;
            return MsVAnchor;
        }

        private LyricsVAlignmentType GetLyricVAlignment(MsoVerticalAnchor msVAnchor)
        {
            LyricsVAlignmentType LVAlign;
            if (msVAnchor == MsoVerticalAnchor.msoAnchorTop)
                LVAlign = LyricsVAlignmentType.Top;
            else if (msVAnchor == MsoVerticalAnchor.msoAnchorMiddle)
                LVAlign = LyricsVAlignmentType.Center;
            else
                LVAlign = LyricsVAlignmentType.Bottom;
            return LVAlign;
        }

        private TextEmphasis GetTextEmphasis(bool bold, bool italic, bool underline)
        {
            TextEmphasis emphasis = TextEmphasis.None;
            if (bold == true)
            {
                emphasis |= TextEmphasis.Bold;
            }
            if (italic == true)
            {
                emphasis |= TextEmphasis.Italic;
            }
            if (underline == true)
            {
                emphasis |= TextEmphasis.UnderLine;
            }
            return emphasis;
        }

        private void SetEmphasis(TextEmphasis emphasis)
        {
            IsBold = ((emphasis & TextEmphasis.Bold) == TextEmphasis.Bold);
            IsItalic = ((emphasis & TextEmphasis.Italic) == TextEmphasis.Italic);
            IsUnderline = ((emphasis & TextEmphasis.UnderLine) == TextEmphasis.UnderLine);
        }

        private int GetColorInteger
            (string color) => int.Parse(color.Remove(0, 3), NumberStyles.HexNumber);

        private string GetColorString
            (int color) => "#FF" + Convert.ToString(color, 16);
       
        private bool IsDuplicatedName(string target, int index)
        {
            bool isDuplicated = false;
            for(int i=0; i < PresetList.Count; i++)
            {
                if (i == index) continue;
                if (target == PresetList[i].Name)
                {
                    isDuplicated = true;
                    break;
                }
            }
            
            return isDuplicated;
        }

        [RelayCommand]
        public void SavePreset()
        {
            if (CurrentSelectedPresetIndex == -1)
            {
                NewPresetDialogResult = PresetDialogResult.Selected;
                PresetGenMethod = PresetGenerationMethod.CopyFromCurrent;
                NewPresetName = string.Empty;
                PresetNameDialogOpen = true;
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("현재 설정을 ");
                sb.Append(PresetList[CurrentSelectedPresetIndex].Name);
                sb.Append("에 저장합니다.");
                if (MessageBox.Show(sb.ToString(), "Save Preset", MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK)
                {
                    PresetList[CurrentSelectedPresetIndex].Options.CopyProperty(GetSlideOptionsFromCurrentSetting());
                    UpdatePresetFile();
                }
                else return;
            }
        }

        [RelayCommand]
        public void CopySlide()
        {   
            PreviewPopupOpen = false;
            var phrases = PresentationMaker.SplitLyrics(LyricsboxText);

            var options = new LyricSlideOptions
            {
                SizeType = SizeType,
                SlideWidth = SlideWidth,
                SlideHeight = SlideHeight,
                BackgroundColor = GetColorInteger(BackgroundColor),
                FontColor = GetColorInteger(ForegroundColor),
                Emphasis = GetTextEmphasis(IsBold,IsItalic,IsUnderline),
                FontName = FontList[FontSelectedIndex],
                FontSize = FontSize,
                Offset = AlignmentOffset.Value,
                VerticalAlignment = GetMsVAnchor(LyricsVAlignment)
            };
            PresentationMaker.MakeLyricsSlides(phrases, options, true);
        }

        [RelayCommand]
        public void Reset()
        {
            CurrentSelectedPresetIndex = -1;
            FontPreviewText = "AaBbCc가나다라1234";
            PreviewPopupOpen = false;
        }

        [RelayCommand]
        public void ChangePreviewStatus()
        {
            PreviewPopupOpen = !PreviewPopupOpen;
        }
    }
}
