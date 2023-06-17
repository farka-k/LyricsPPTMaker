using LyricsPPTMaker.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace LyricsPPTMaker
{
    public class PresetXMLSerializer
    {
        XmlSerializer serializer;
        private ObservableCollection<Preset>? _presets;
        private bool isTargetSet = false;
        public PresetXMLSerializer()
        {
            serializer=new XmlSerializer(typeof(ObservableCollection<Preset>));
        }

        public void SetPresetList(ObservableCollection<Preset> presets)
        {
            _presets = presets;
            isTargetSet = true;
        }
        public void Serialize()
        {
            if (!isTargetSet) { return; }
            using (StreamWriter wr=new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"\Presets.xml"))
            {
                serializer.Serialize(wr, _presets);
            }
        }

        public ObservableCollection<Preset>? Deserialize(string outpath)
        {
            using (var reader= new StreamReader(outpath))
            {
                _presets = (ObservableCollection<Preset>?) serializer.Deserialize(reader);
            }
            isTargetSet = true;
            return _presets;
        }
    }

}
