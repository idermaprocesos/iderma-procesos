using System.Collections.ObjectModel;
using Iderma_Procesos.Models;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace Iderma_Procesos;

public sealed partial class CombinarDialog : ContentDialog
{
    private readonly ObservableCollection<EntradaCombinar> _entradas = new();

    public CombinarDialog()
    {
        InitializeComponent();
        Lista.ItemsSource = _entradas;
        PrimaryButtonClick += AlConfirmar;
        SecondaryButtonClick += AlConfirmar;
    }

    public IReadOnlyList<Procedimiento> Seleccion { get; private set; } = [];

    public void Cargar(IEnumerable<Procedimiento> procedimientos)
    {
        _entradas.Clear();
        foreach (var procedimiento in procedimientos)
            _entradas.Add(new EntradaCombinar(procedimiento));
    }

    private void AlConfirmar(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        var elegidos = _entradas.Where(e => e.Incluido).Select(e => e.Procedimiento).ToList();
        if (elegidos.Count == 0)
        {
            args.Cancel = true;
            Mensaje.Visibility = Visibility.Visible;
            return;
        }

        Mensaje.Visibility = Visibility.Collapsed;
        Seleccion = elegidos;
    }

    private void Subir_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button boton || boton.Tag is not EntradaCombinar entrada)
            return;

        var indice = _entradas.IndexOf(entrada);
        if (indice > 0)
            _entradas.Move(indice, indice - 1);
    }

    private void Bajar_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button boton || boton.Tag is not EntradaCombinar entrada)
            return;

        var indice = _entradas.IndexOf(entrada);
        if (indice >= 0 && indice < _entradas.Count - 1)
            _entradas.Move(indice, indice + 1);
    }
}

public sealed class EntradaCombinar
{
    public EntradaCombinar(Procedimiento procedimiento) => Procedimiento = procedimiento;

    public Procedimiento Procedimiento { get; }

    public bool Incluido { get; set; } = true;

    public string Codigo => Procedimiento.CodigoLista;

    public string Titulo => Procedimiento.TituloLista;
}
