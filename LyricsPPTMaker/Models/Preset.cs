using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LyricsPPTMaker.Models
{
    [Serializable]
    public class Preset
    {
        public string Name { get; set; }
        public LyricSlideOptions Options { get; set; }
        public Preset() {
            Name = String.Empty;
            Options = new LyricSlideOptions();
        }
        public Preset(string name, LyricSlideOptions options)
        {
            Name = name;
            Options = options;
        }
    }
    public enum PresetDialogResult { Selected, Cancel, None };
    public enum PresetGenerationMethod { Default, CopyFromCurrent, None };
}
