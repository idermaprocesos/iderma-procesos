using System.Collections.ObjectModel;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Iderma_Procesos.Models;
using Iderma_Procesos.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.System;
using WinRT.Interop;

namespace Iderma_Procesos;

public sealed partial class MainPage : Page
{
    private readonly ProcedimientoStore _almacen = new();
    private readonly DispatcherTimer _guardado = new() { Interval = TimeSpan.FromMilliseconds(350) };
    private readonly HashSet<Procedimiento> _enganchados = new();
    private ObservableCollection<Procedimiento> _procedimientos = new();
    private bool _suspendido;

    public MainPage()
    {
        InitializeComponent();
        _guardado.Tick += (_, _) =>
        {
            _guardado.Stop();
            GuardarAhora();
        };
        Loaded += (_, _) => Cargar();
    }

    private void Cargar()
    {
        if (Lista.ItemsSource is not null)
            return;

        _suspendido = true;
        _procedimientos = _almacen.Cargar();
        foreach (var procedimiento in _procedimientos)
            Enganchar(procedimiento);

        Lista.ItemsSource = _procedimientos;
        if (_procedimientos.Count > 0)
            Lista.SelectedIndex = 0;
        else
            Mostrar(null);

        _suspendido = false;
        if (App.Ventana is not null)
            App.Ventana.Closed += (_, _) => GuardarAhora();
    }

    private void Enganchar(Procedimiento procedimiento)
    {
        if (!_enganchados.Add(procedimiento))
            return;

        procedimiento.PropertyChanged += AlCambiar;
        procedimiento.Secciones.CollectionChanged += SeccionesCambiadas;
        foreach (var seccion in procedimiento.Secciones)
            EngancharSeccion(seccion);
    }

    private void EngancharSeccion(Seccion seccion)
    {
        seccion.AsegurarFilas();
        seccion.PropertyChanged += AlCambiar;
        seccion.FilasDesarrollo.CollectionChanged += FilasCambiadas;
        seccion.FilasCambio.CollectionChanged += FilasCambiadas;
        foreach (var fila in seccion.FilasDesarrollo)
            fila.PropertyChanged += AlCambiar;
        foreach (var fila in seccion.FilasCambio)
            fila.PropertyChanged += AlCambiar;
    }

    private void SeccionesCambiadas(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (Seccion seccion in e.NewItems)
                EngancharSeccion(seccion);
        }

        ProgramarGuardado();
    }

    private void FilasCambiadas(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (Observable fila in e.NewItems)
                fila.PropertyChanged += AlCambiar;
        }

        ProgramarGuardado();
    }

    private void AlCambiar(object? sender, PropertyChangedEventArgs e) => ProgramarGuardado();

    private void ProgramarGuardado()
    {
        if (_suspendido)
            return;

        EstadoGuardado.Text = "Guardando…";
        _guardado.Stop();
        _guardado.Start();
    }

    private void GuardarAhora()
    {
        _guardado.Stop();
        _almacen.Guardar(_procedimientos);
        if (EstadoGuardado is not null)
            EstadoGuardado.Text = "Guardado";
    }

    private void Lista_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        Mostrar(Lista.SelectedItem as Procedimiento);
    }

    private void Mostrar(Procedimiento? procedimiento)
    {
        Editor.DataContext = procedimiento;
        Editor.Visibility = procedimiento is null ? Visibility.Collapsed : Visibility.Visible;
        Vacio.Visibility = procedimiento is null ? Visibility.Visible : Visibility.Collapsed;
    }

    private void Nuevo_Click(object sender, RoutedEventArgs e) => CrearNuevo();

    private void NuevoAtajo(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        CrearNuevo();
        args.Handled = true;
    }

    private void CrearNuevo()
    {
        var procedimiento = Procedimiento.CrearNuevo();
        Enganchar(procedimiento);
        _procedimientos.Insert(0, procedimiento);
        Lista.SelectedItem = procedimiento;
        GuardarAhora();
    }

    private void Duplicar_Click(object sender, RoutedEventArgs e)
    {
        if (Editor.DataContext is not Procedimiento actual)
            return;

        var copia = actual.Copiar();
        Enganchar(copia);
        var indice = _procedimientos.IndexOf(actual);
        _procedimientos.Insert(indice + 1, copia);
        Lista.SelectedItem = copia;
        GuardarAhora();
    }

    private async void Eliminar_Click(object sender, RoutedEventArgs e)
    {
        if (Editor.DataContext is not Procedimiento actual)
            return;

        var dialogo = new ContentDialog
        {
            Title = "Eliminar procedimiento",
            Content = string.IsNullOrWhiteSpace(actual.Titulo)
                ? "Se quitará este procedimiento de la lista."
                : $"Se quitará «{actual.Titulo}» de la lista.",
            PrimaryButtonText = "Eliminar",
            CloseButtonText = "Cancelar",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = XamlRoot
        };

        if (await dialogo.ShowAsync() != ContentDialogResult.Primary)
            return;

        var indice = _procedimientos.IndexOf(actual);
        _procedimientos.Remove(actual);
        if (_procedimientos.Count == 0)
            Mostrar(null);
        else
            Lista.SelectedIndex = Math.Min(indice, _procedimientos.Count - 1);

        GuardarAhora();
    }

    private void AgregarDesarrollo_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button boton || boton.Tag is not Seccion seccion)
            return;

        seccion.AsegurarFilas();
        seccion.FilasDesarrollo.Add(new FilaDesarrollo());
    }

    private void QuitarDesarrollo_Click(object sender, RoutedEventArgs e)
    {
        if (Editor.DataContext is not Procedimiento actual || sender is not Button boton || boton.Tag is not FilaDesarrollo fila)
            return;

        foreach (var seccion in actual.Secciones)
        {
            if (seccion.FilasDesarrollo.Remove(fila))
                return;
        }
    }

    private void AgregarCambio_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button boton || boton.Tag is not Seccion seccion)
            return;

        seccion.AsegurarFilas();
        seccion.FilasCambio.Add(new FilaCambio());
    }

    private void QuitarCambio_Click(object sender, RoutedEventArgs e)
    {
        if (Editor.DataContext is not Procedimiento actual || sender is not Button boton || boton.Tag is not FilaCambio fila)
            return;

        foreach (var seccion in actual.Secciones)
        {
            if (seccion.FilasCambio.Remove(fila))
                return;
        }
    }

    private async void Generar_Click(object sender, RoutedEventArgs e) => await Generar(pdf: false);

    private async void GenerarPdf_Click(object sender, RoutedEventArgs e) => await Generar(pdf: true);

    private async void Combinar_Click(object sender, RoutedEventArgs e)
    {
        if (_procedimientos.Count == 0)
        {
            MostrarAviso("Cree al menos un procedimiento antes de combinar.", InfoBarSeverity.Warning);
            return;
        }

        var dialogo = new CombinarDialog { XamlRoot = XamlRoot };
        dialogo.Cargar(_procedimientos);
        var resultado = await dialogo.ShowAsync();
        if (resultado == ContentDialogResult.None)
            return;

        await Generar(pdf: resultado == ContentDialogResult.Secondary, procedimientos: dialogo.Seleccion);
    }

    private async Task Generar(bool pdf, IReadOnlyList<Procedimiento>? procedimientos = null)
    {
        List<Procedimiento> lista;
        if (procedimientos is not null)
            lista = procedimientos.ToList();
        else if (Editor.DataContext is Procedimiento actual)
            lista = [actual];
        else
            return;

        if (lista.Count == 0)
        {
            MostrarAviso("Seleccione al menos un procedimiento.", InfoBarSeverity.Warning);
            return;
        }

        if (lista.Any(p => string.IsNullOrWhiteSpace(p.Titulo) || string.IsNullOrWhiteSpace(p.Codigo)))
        {
            MostrarAviso("Escriba el título y el código de cada procedimiento antes de generar el documento.", InfoBarSeverity.Warning);
            return;
        }

        GuardarAhora();

        try
        {
            if (App.Ventana is null)
                return;

            var selector = new FileSavePicker();
            InitializeWithWindow.Initialize(selector, WindowNative.GetWindowHandle(App.Ventana));
            selector.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            selector.FileTypeChoices.Add(pdf ? "Documento PDF" : "Documento Word", new List<string> { pdf ? ".pdf" : ".docx" });
            selector.SuggestedFileName = NombreArchivo(lista);

            var archivo = await selector.PickSaveFileAsync();
            if (archivo is null)
                return;

            var logo = await LeerLogoAsync();
            if (pdf)
            {
                MostrarAviso("Creando el PDF…", InfoBarSeverity.Informational);
                var docx = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
                try
                {
                    await using (var stream = File.Create(docx))
                        WordDocumentService.CrearCombinado(stream, lista, logo);

                    var destino = archivo.Path;
                    var copiar = string.IsNullOrWhiteSpace(destino) || !Path.IsPathRooted(destino);
                    var pdfTemporal = copiar ? Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf") : destino;
                    await ExportarPdfAsync(docx, pdfTemporal);
                    if (copiar)
                    {
                        await using var lectura = File.OpenRead(pdfTemporal);
                        await using var escritura = await archivo.OpenStreamForWriteAsync();
                        escritura.SetLength(0);
                        await lectura.CopyToAsync(escritura);
                        File.Delete(pdfTemporal);
                    }
                }
                finally
                {
                    if (File.Exists(docx))
                        File.Delete(docx);
                }

                MostrarAviso("PDF creado.", InfoBarSeverity.Success);
            }
            else
            {
                using (var stream = await archivo.OpenStreamForWriteAsync())
                {
                    stream.SetLength(0);
                    WordDocumentService.CrearCombinado(stream, lista, logo);
                }

                MostrarAviso("Documento creado. Si Word pregunta por los campos, elija Sí para numerar las páginas.", InfoBarSeverity.Success);
            }

            await Launcher.LaunchFileAsync(archivo);
        }
        catch (IOException)
        {
            MostrarAviso("No se pudo guardar. Cierre el archivo si ya está abierto.", InfoBarSeverity.Error);
        }
        catch (COMException)
        {
            MostrarAviso("No se pudo crear el PDF. Cierre el archivo si ya está abierto y compruebe que Word esté instalado.", InfoBarSeverity.Error);
        }
        catch (Exception ex)
        {
            MostrarAviso("No se pudo crear el documento. " + ex.Message, InfoBarSeverity.Error);
        }
    }

    private static Task ExportarPdfAsync(string docx, string pdf)
    {
        var origen = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var hilo = new Thread(() =>
        {
            try
            {
                WordDocumentService.ExportarPdf(docx, pdf, () => origen.TrySetResult());
            }
            catch (Exception ex)
            {
                origen.TrySetException(ex);
            }
        });
        hilo.SetApartmentState(ApartmentState.STA);
        hilo.IsBackground = true;
        hilo.Start();
        return origen.Task;
    }

    private static async Task<byte[]> LeerLogoAsync()
    {
        var local = Path.Combine(AppContext.BaseDirectory, "Assets", "logo.png");
        if (File.Exists(local))
            return await File.ReadAllBytesAsync(local);

        var archivo = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///Assets/logo.png"));
        var buffer = await FileIO.ReadBufferAsync(archivo);
        return buffer.ToArray();
    }

    private static string NombreArchivo(IReadOnlyList<Procedimiento> procedimientos)
    {
        var nombre = procedimientos.Count == 1
            ? $"{procedimientos[0].Codigo} {procedimientos[0].Titulo}".Trim()
            : string.Join(" ", procedimientos.Select(p => p.Codigo).Where(c => !string.IsNullOrWhiteSpace(c)));

        if (procedimientos.Count > 1 && string.IsNullOrWhiteSpace(nombre))
            nombre = "Procedimientos combinados";

        foreach (var caracter in Path.GetInvalidFileNameChars())
            nombre = nombre.Replace(caracter, ' ');

        nombre = string.Join(' ', nombre.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        if (nombre.Length > 90)
            nombre = nombre[..90].Trim();

        return string.IsNullOrWhiteSpace(nombre) ? "Procedimiento" : nombre;
    }

    private void MostrarAviso(string mensaje, InfoBarSeverity severidad)
    {
        Aviso.Message = mensaje;
        Aviso.Severity = severidad;
        Aviso.IsOpen = true;
    }
}
