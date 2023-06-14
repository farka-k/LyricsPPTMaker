using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Wpf.AvalonDock.Controls;

namespace LyricsPPTMaker.Models
{
    public class SongInfo
    {
        public string Title { get; set; }
        public string Artist { get; set; }
        public string Album { get; set; }
        public string Lyrics { get; set; }
        
        public SongInfo(string title, string artist, string album, string lyrics)
        {
            Title = title;
            Artist = artist;
            Album = album;
            Lyrics = lyrics;
        }
    }
}
