using System.Collections.ObjectModel;
using System.Text.Json;
using Iderma_Procesos.Models;

namespace Iderma_Procesos.Services;

public sealed class ProcedimientoStore
{
    private static readonly JsonSerializerOptions Opciones = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _ruta;

    public ProcedimientoStore()
    {
        var carpeta = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "IdermaCapilar",
            "Procedimientos");
        Directory.CreateDirectory(carpeta);
        _ruta = Path.Combine(carpeta, "procedimientos.json");
    }

    public ObservableCollection<Procedimiento> Cargar()
    {
        if (!File.Exists(_ruta))
            return new ObservableCollection<Procedimiento> { Procedimiento.Ejemplo() };

        try
        {
            var json = File.ReadAllText(_ruta);
            var lista = JsonSerializer.Deserialize<List<Procedimiento>>(json, Opciones);
            if (lista is null || lista.Count == 0)
                return new ObservableCollection<Procedimiento> { Procedimiento.Ejemplo() };

            foreach (var procedimiento in lista)
            {
                procedimiento.Normalizar();
                procedimiento.UsarHoySiEsFechaDeMuestra();
            }

            return new ObservableCollection<Procedimiento>(lista);
        }
        catch
        {
            return new ObservableCollection<Procedimiento> { Procedimiento.Ejemplo() };
        }
    }

    public void Guardar(IEnumerable<Procedimiento> procedimientos)
    {
        var json = JsonSerializer.Serialize(procedimientos, Opciones);
        File.WriteAllText(_ruta, json);
    }
}
