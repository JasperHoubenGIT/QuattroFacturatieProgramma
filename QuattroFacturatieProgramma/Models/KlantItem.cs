using CommunityToolkit.Mvvm.ComponentModel;

namespace QuattroFacturatieProgramma.Models;

public partial class KlantItem : ObservableObject
{
    [ObservableProperty]
    private string _naam = string.Empty;

    [ObservableProperty]
    private double _bedrag;

    [ObservableProperty]
    private bool _isGeselecteerd;

    public string BedragFormatted => $"€ {Bedrag:F2}";
}