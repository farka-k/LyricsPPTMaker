using System;
using HtmlAgilityPack;
using LyricsPPTMaker.Models;

namespace LyricsPPTMaker
{
    public static class LyricsCollector
    {
        public static SongInfo GetLyrics(string url)
        {
            string songTitle = String.Empty;
            string artist = String.Empty;
            string albumTitle= String.Empty;
            string fullLyrics = String.Empty;
            HtmlWeb web = new HtmlWeb();
            HtmlDocument htmlDoc = web.Load(url);

            var contentBody = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='hyrendContentBody']");
            var article = contentBody.SelectSingleNode("article");
            songTitle = article.SelectSingleNode("header/div/h1").InnerText.Trim();
            
            var summary = article.SelectSingleNode("section[@class='sectionPadding summaryInfo summaryTrack']");
            var infoTable = summary.SelectSingleNode("div/div[@class='basicInfo']/table");
            var artistColumn = infoTable.SelectSingleNode("tbody/tr[th='아티스트']/td");
            if (artistColumn.SelectSingleNode("a") != null)
                artist = artistColumn.SelectSingleNode("a").InnerText.Trim();
            else
                artist = artistColumn.InnerText.Trim();
            var albumInfo = infoTable.SelectSingleNode("tbody/tr[th='앨범']/td/a");
            if (albumInfo!=null)
                albumTitle = albumInfo.InnerText;

            var lyricsSection = article.SelectSingleNode("section[@class='sectionPadding contents lyrics']");
            var lyricsContainer = lyricsSection.SelectSingleNode("div/div[@class='lyricsContainer']");
            var lyrcisOuter = lyricsContainer.SelectSingleNode("p/xmp");
            if (lyrcisOuter != null)
                fullLyrics = lyrcisOuter.InnerText;
            else fullLyrics = "가사가 없습니다";
            SongInfo songInfo = new SongInfo(songTitle, artist, albumTitle, fullLyrics);
            return songInfo;
        }
    }
}
