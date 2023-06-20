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
        private string docFolder;
        public string presetFolderName;
        public string presetFolderPath;
        public string presetFilePath;
        public PresetXMLSerializer()
        {
            serializer=new XmlSerializer(typeof(ObservableCollection<Preset>));
            docFolder= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            presetFolderName = "LyricsPPTMaker";
            presetFolderPath = docFolder + @"\" + presetFolderName;
            presetFilePath = presetFolderPath + @"\Presets.xml";
        }

        public void SetPresetList(ObservableCollection<Preset> presets)
        {
            _presets = presets;
            isTargetSet = true;
        }
        public void Serialize()
        {
            if (!isTargetSet) { return; }
            CheckPresetDirectory();
            using (StreamWriter wr=new StreamWriter(presetFilePath))
            {
                serializer.Serialize(wr, _presets);
            }
        }

        public ObservableCollection<Preset> Deserialize()
        {
            using (var reader= new StreamReader(presetFilePath))
            {
                _presets = (ObservableCollection<Preset>?) serializer.Deserialize(reader);
            }
            isTargetSet = true;
            return _presets;
        }

        public void CheckPresetDirectory()
        {
            if (!Directory.Exists(docFolder + @"\LyricsPPTMaker")) Directory.CreateDirectory(presetFolderPath);
        }

        public bool PresetFileCreated() => File.Exists(presetFilePath);
        
    }

}
