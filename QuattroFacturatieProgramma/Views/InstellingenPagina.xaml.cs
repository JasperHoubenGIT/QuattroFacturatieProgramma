using Microsoft.Maui.Storage;

namespace QuattroFacturatieProgramma.Views;

public partial class InstellingenPagina : ContentPage
{
    public InstellingenPagina()
    {
        InitializeComponent();
        PadEntry.Text = Preferences.Get("FactuurOpslagPad", "");
    }

    private async void Opslaan_Clicked(object sender, EventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(PadEntry.Text) && Directory.Exists(PadEntry.Text))
        {
            Preferences.Set("FactuurOpslagPad", PadEntry.Text);
            await DisplayAlert("Opgeslagen", "Opslagpad is opgeslagen.", "OK");
        }
        else
        {
            await DisplayAlert("Ongeldig pad", "Voer een geldig bestaand pad in.", "OK");
        }
    }
}
