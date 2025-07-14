using QuattroFacturatieProgramma.ViewModels;

namespace QuattroFacturatieProgramma;

public partial class MainPage : ContentPage
{
    public MainPage(MainViewModel viewModel)
    {
        InitializeComponent();
        BindingContext = viewModel;
    }
}
