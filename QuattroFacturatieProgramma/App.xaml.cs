using QuattroFacturatieProgramma.Views;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Maui.Controls.Application;

namespace QuattroFacturatieProgramma;
using QuattroFacturatieProgramma.Views;

public partial class App : Application
{
    public App()
    {
        InitializeComponent();
        MainPage = new AppShell(); // NIET null
        Routing.RegisterRoute(nameof(InstellingenPagina), typeof(InstellingenPagina));
    }

}
