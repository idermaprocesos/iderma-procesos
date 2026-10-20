using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;

namespace Iderma_Procesos;

public sealed partial class MainWindow : Window
{
    public MainWindow()
    {
        App.Ventana = this;
        InitializeComponent();
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);
        AppWindow.SetIcon("Assets/AppIcon.ico");

        if (AppWindow.Presenter is OverlappedPresenter presenter)
            presenter.Maximize();

        RootFrame.Navigate(typeof(MainPage));
    }
}
