#pragma warning disable IDE1006 // 명명 스타일
using System.Collections.Generic;

namespace LyricsPPTMaker.Models
{
    public class GoogleCustomSearchAPIResponseModel
    {
        public List<SearchResultItem>? items { get; set; }

    }

    public class SearchResultItem
    {
        public string? title { get; set; }
        public string? link { get; set; }
    }
}
#pragma warning restore IDE1006 // 명명 스타일
