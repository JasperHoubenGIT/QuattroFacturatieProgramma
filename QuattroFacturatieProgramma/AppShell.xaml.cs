namespace QuattroFacturatieProgramma;
using QuattroFacturatieProgramma.Views;

public partial class AppShell : Shell
{
    public AppShell()
    {
        InitializeComponent();
        Routing.RegisterRoute(nameof(InstellingenPagina), typeof(InstellingenPagina));
    }
}