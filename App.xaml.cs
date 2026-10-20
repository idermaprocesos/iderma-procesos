using Microsoft.UI.Xaml;

namespace Iderma_Procesos;

public partial class App : Application
{
    public static Window? Ventana { get; set; }

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var ventana = new MainWindow();
        Ventana = ventana;
        ventana.Activate();
    }
}
