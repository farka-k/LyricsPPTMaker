using CommunityToolkit.Mvvm.ComponentModel;
namespace LyricsPPTMaker.Bases
{
    /// <summary>
    /// ViewModel Base
    /// </summary>
    public partial class ViewModelBase : ObservableObject
    {
        [ObservableProperty]
        private string? _title;
    }
}
