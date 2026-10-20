using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace Iderma_Procesos.Models;

public abstract class Observable : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? nombre = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;

        field = value;
        Notificar(nombre);
        return true;
    }

    protected void Notificar(string? nombre) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nombre));
}

public sealed class FilaDesarrollo : Observable
{
    private string _area = "";
    private string _responsable = "";
    private string _actividad = "";

    public string Area
    {
        get => _area;
        set => Set(ref _area, value);
    }

    public string Responsable
    {
        get => _responsable;
        set => Set(ref _responsable, value);
    }

    public string Actividad
    {
        get => _actividad;
        set => Set(ref _actividad, value);
    }

    public FilaDesarrollo Clonar() => new()
    {
        Area = Area,
        Responsable = Responsable,
        Actividad = Actividad
    };
}

public sealed class FilaCambio : Observable
{
    private string _version = "";
    private string _fecha = "";
    private string _descripcion = "";
    private string _elaboro = "";
    private string _reviso = "";
    private string _aprobo = "";

    public string Version
    {
        get => _version;
        set => Set(ref _version, value);
    }

    public string Fecha
    {
        get => _fecha;
        set => Set(ref _fecha, value);
    }

    public string Descripcion
    {
        get => _descripcion;
        set => Set(ref _descripcion, value);
    }

    public string Elaboro
    {
        get => _elaboro;
        set => Set(ref _elaboro, value);
    }

    public string Reviso
    {
        get => _reviso;
        set => Set(ref _reviso, value);
    }

    public string Aprobo
    {
        get => _aprobo;
        set => Set(ref _aprobo, value);
    }

    public FilaCambio Clonar() => new()
    {
        Version = Version,
        Fecha = Fecha,
        Descripcion = Descripcion,
        Elaboro = Elaboro,
        Reviso = Reviso,
        Aprobo = Aprobo
    };
}

public sealed class Seccion : Observable
{
    public const string TipoTexto = "texto";
    public const string TipoDesarrollo = "desarrollo";
    public const string TipoCambios = "cambios";

    private string _titulo = "";
    private string _contenido = "";
    private string _tipo = TipoTexto;

    public string Titulo
    {
        get => _titulo;
        set => Set(ref _titulo, value);
    }

    public string Contenido
    {
        get => _contenido;
        set => Set(ref _contenido, value);
    }

    public string Tipo
    {
        get => _tipo;
        set => Set(ref _tipo, value);
    }

    public ObservableCollection<FilaDesarrollo> FilasDesarrollo { get; set; } = new();

    public ObservableCollection<FilaCambio> FilasCambio { get; set; } = new();

    [JsonIgnore]
    public string Ayuda { get; set; } = "";

    public void AsegurarFilas()
    {
        FilasDesarrollo ??= new ObservableCollection<FilaDesarrollo>();
        FilasCambio ??= new ObservableCollection<FilaCambio>();
    }

    public Seccion Clonar()
    {
        AsegurarFilas();
        var copia = new Seccion
        {
            Titulo = Titulo,
            Contenido = Contenido,
            Tipo = string.IsNullOrWhiteSpace(Tipo) ? TipoTexto : Tipo,
            Ayuda = Ayuda
        };
        foreach (var fila in FilasDesarrollo)
            copia.FilasDesarrollo.Add(fila.Clonar());
        foreach (var fila in FilasCambio)
            copia.FilasCambio.Add(fila.Clonar());
        return copia;
    }
}

public sealed class Procedimiento : Observable
{
    private static readonly CultureInfo Cultura = CultureInfo.GetCultureInfo("es-PE");

    private string _titulo = "";
    private string _codigo = "";
    private string _version = "1";
    private DateTime _fechaDocumento = DateTime.Today;
    private DateTime _fechaElaboracion = DateTime.Today;
    private DateTime _fechaRevision = DateTime.Today;
    private DateTime _fechaAprobacion = DateTime.Today;
    private string _elaboradoNombre = "Piero Zavala Ayala";
    private string _elaboradoCargo = "Encargado de Calidad y Procesos";
    private string _revisadoNombre = "Maribel Dávila Zamalloa";
    private string _revisadoCargo = "Gerente General";
    private string _aprobadoNombre = "Maribel Dávila Zamalloa";
    private string _aprobadoCargo = "Gerente General";

    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Titulo
    {
        get => _titulo;
        set
        {
            if (Set(ref _titulo, value))
            {
                Notificar(nameof(TituloMayusculas));
                Notificar(nameof(TituloLista));
            }
        }
    }

    public string Codigo
    {
        get => _codigo;
        set
        {
            if (Set(ref _codigo, value))
                Notificar(nameof(CodigoLista));
        }
    }

    public string Version
    {
        get => _version;
        set => Set(ref _version, value);
    }

    public DateTime FechaDocumento
    {
        get => _fechaDocumento;
        set
        {
            if (Set(ref _fechaDocumento, value.Date))
            {
                Notificar(nameof(FechaDocumentoTexto));
                Notificar(nameof(FechaDocumentoPicker));
            }
        }
    }

    public DateTime FechaElaboracion
    {
        get => _fechaElaboracion;
        set
        {
            if (Set(ref _fechaElaboracion, value.Date))
            {
                Notificar(nameof(FechaElaboracionTexto));
                Notificar(nameof(FechaElaboracionLinea));
                Notificar(nameof(FechaElaboracionPicker));
            }
        }
    }

    public DateTime FechaRevision
    {
        get => _fechaRevision;
        set
        {
            if (Set(ref _fechaRevision, value.Date))
            {
                Notificar(nameof(FechaRevisionTexto));
                Notificar(nameof(FechaRevisionLinea));
                Notificar(nameof(FechaRevisionPicker));
            }
        }
    }

    public DateTime FechaAprobacion
    {
        get => _fechaAprobacion;
        set
        {
            if (Set(ref _fechaAprobacion, value.Date))
            {
                Notificar(nameof(FechaAprobacionTexto));
                Notificar(nameof(FechaAprobacionLinea));
                Notificar(nameof(FechaAprobacionPicker));
            }
        }
    }

    public string ElaboradoNombre
    {
        get => _elaboradoNombre;
        set => Set(ref _elaboradoNombre, value);
    }

    public string ElaboradoCargo
    {
        get => _elaboradoCargo;
        set => Set(ref _elaboradoCargo, value);
    }

    public string RevisadoNombre
    {
        get => _revisadoNombre;
        set => Set(ref _revisadoNombre, value);
    }

    public string RevisadoCargo
    {
        get => _revisadoCargo;
        set => Set(ref _revisadoCargo, value);
    }

    public string AprobadoNombre
    {
        get => _aprobadoNombre;
        set => Set(ref _aprobadoNombre, value);
    }

    public string AprobadoCargo
    {
        get => _aprobadoCargo;
        set => Set(ref _aprobadoCargo, value);
    }

    public ObservableCollection<Seccion> Secciones { get; set; } = new();

    [JsonIgnore]
    public string TituloMayusculas => (Titulo ?? "").ToUpper(Cultura);

    [JsonIgnore]
    public string TituloLista => string.IsNullOrWhiteSpace(Titulo) ? "Sin título" : Titulo;

    [JsonIgnore]
    public string CodigoLista => string.IsNullOrWhiteSpace(Codigo) ? "Sin código" : Codigo;

    [JsonIgnore]
    public string FechaDocumentoTexto => FechaDocumento.ToString("dd/MM/yyyy", Cultura);

    [JsonIgnore]
    public string FechaElaboracionTexto => FechaElaboracion.ToString("dd/MM/yy", Cultura);

    [JsonIgnore]
    public string FechaRevisionTexto => FechaRevision.ToString("dd/MM/yy", Cultura);

    [JsonIgnore]
    public string FechaAprobacionTexto => FechaAprobacion.ToString("dd/MM/yy", Cultura);

    [JsonIgnore]
    public string FechaElaboracionLinea => "Fecha: " + FechaElaboracionTexto;

    [JsonIgnore]
    public string FechaRevisionLinea => "Fecha: " + FechaRevisionTexto;

    [JsonIgnore]
    public string FechaAprobacionLinea => "Fecha: " + FechaAprobacionTexto;

    [JsonIgnore]
    public DateTimeOffset? FechaDocumentoPicker
    {
        get => Fecha(FechaDocumento);
        set { if (value is not null) FechaDocumento = value.Value.DateTime; }
    }

    [JsonIgnore]
    public DateTimeOffset? FechaElaboracionPicker
    {
        get => Fecha(FechaElaboracion);
        set { if (value is not null) FechaElaboracion = value.Value.DateTime; }
    }

    [JsonIgnore]
    public DateTimeOffset? FechaRevisionPicker
    {
        get => Fecha(FechaRevision);
        set { if (value is not null) FechaRevision = value.Value.DateTime; }
    }

    [JsonIgnore]
    public DateTimeOffset? FechaAprobacionPicker
    {
        get => Fecha(FechaAprobacion);
        set { if (value is not null) FechaAprobacion = value.Value.DateTime; }
    }

    public void UsarHoySiEsFechaDeMuestra()
    {
        var hoy = DateTime.Today;
        if (FechaDocumento.Date == new DateTime(2026, 10, 2))
            FechaDocumento = hoy;
        if (FechaElaboracion.Date == new DateTime(2025, 10, 1))
            FechaElaboracion = hoy;
        if (FechaRevision.Date == new DateTime(2025, 10, 1))
            FechaRevision = hoy;
        if (FechaAprobacion.Date == new DateTime(2025, 10, 1))
            FechaAprobacion = hoy;
    }

    public static Procedimiento Ejemplo()
    {
        var procedimiento = CrearBase();
        procedimiento.Titulo = "Procedimiento de ordenamiento de carpetas del proceso contable";
        procedimiento.Codigo = "PRI-CTB-07";
        procedimiento.Version = "1";
        var hoy = DateTime.Today;
        procedimiento.FechaDocumento = hoy;
        procedimiento.FechaElaboracion = hoy;
        procedimiento.FechaRevision = hoy;
        procedimiento.FechaAprobacion = hoy;
        return procedimiento;
    }

    public static Procedimiento CrearNuevo()
    {
        var procedimiento = CrearBase();
        procedimiento.Version = "1";
        var hoy = DateTime.Today;
        procedimiento.FechaDocumento = hoy;
        procedimiento.FechaElaboracion = hoy;
        procedimiento.FechaRevision = hoy;
        procedimiento.FechaAprobacion = hoy;
        return procedimiento;
    }

    public Procedimiento Copiar()
    {
        var copia = CrearNuevo();
        copia.Titulo = string.IsNullOrWhiteSpace(Titulo) ? "" : Titulo + " (copia)";
        copia.Codigo = Codigo;
        copia.Version = Version;
        copia.FechaDocumento = FechaDocumento;
        copia.FechaElaboracion = FechaElaboracion;
        copia.FechaRevision = FechaRevision;
        copia.FechaAprobacion = FechaAprobacion;
        copia.ElaboradoNombre = ElaboradoNombre;
        copia.ElaboradoCargo = ElaboradoCargo;
        copia.RevisadoNombre = RevisadoNombre;
        copia.RevisadoCargo = RevisadoCargo;
        copia.AprobadoNombre = AprobadoNombre;
        copia.AprobadoCargo = AprobadoCargo;
        copia.Secciones = new ObservableCollection<Seccion>(Secciones.Select(s => s.Clonar()));
        return copia;
    }

    public void Normalizar()
    {
        Secciones ??= new ObservableCollection<Seccion>();
        var plantilla = Plantilla();
        var esperadas = plantilla.Select(s => s.Titulo).ToArray();
        var actuales = Secciones.Select(s => (s.Titulo ?? "").Trim()).ToArray();
        var coincide = actuales.Length == esperadas.Length
            && actuales.Zip(esperadas).All(par => string.Equals(par.First, par.Second, StringComparison.OrdinalIgnoreCase));

        if (!coincide)
        {
            Secciones = new ObservableCollection<Seccion>(plantilla);
            return;
        }

        for (var i = 0; i < Secciones.Count; i++)
        {
            Secciones[i].AsegurarFilas();
            Secciones[i].Tipo = plantilla[i].Tipo;
            Secciones[i].Ayuda = plantilla[i].Ayuda;
            Secciones[i].Titulo = plantilla[i].Titulo;
        }
    }

    private static Procedimiento CrearBase()
    {
        return new Procedimiento
        {
            ElaboradoNombre = "Piero Zavala Ayala",
            ElaboradoCargo = "Encargado de Calidad y Procesos",
            RevisadoNombre = "Maribel Dávila Zamalloa",
            RevisadoCargo = "Gerente General",
            AprobadoNombre = "Maribel Dávila Zamalloa",
            AprobadoCargo = "Gerente General",
            Secciones = new ObservableCollection<Seccion>(Plantilla())
        };
    }

    private static List<Seccion> Plantilla()
    {
        return
        [
            new Seccion
            {
                Titulo = "Objetivos",
                Ayuda = "Indique para qué se hace este procedimiento."
            },
            new Seccion
            {
                Titulo = "Alcance",
                Ayuda = "Indique a quién o a qué aplica."
            },
            new Seccion
            {
                Titulo = "Referencias normativas",
                Contenido = "Norma Internacional ISO 9001:2015 Sistemas de gestión de la calidad",
                Ayuda = "Normas, leyes u otros documentos que se citan."
            },
            new Seccion
            {
                Titulo = "Responsables",
                Ayuda = "Indique quién interviene y de qué se encarga."
            },
            new Seccion
            {
                Titulo = "Definiciones y abreviaturas",
                Ayuda = "Escriba los términos y las siglas que se usan."
            },
            new Seccion
            {
                Titulo = "Desarrollo del procedimiento",
                Tipo = Seccion.TipoDesarrollo,
                Ayuda = "Cada fila es un paso: área, responsable y la actividad detallada.",
                FilasDesarrollo = new ObservableCollection<FilaDesarrollo> { new() }
            },
            new Seccion
            {
                Titulo = "Formatos asociados",
                Ayuda = "Indique los formatos o registros que se usan."
            },
            new Seccion
            {
                Titulo = "Anexos",
                Ayuda = "Indique los anexos. Si no hay, escriba que no aplica."
            },
            new Seccion
            {
                Titulo = "Diagrama de flujo",
                Ayuda = "Describa el diagrama o indique en qué anexo está."
            },
            new Seccion
            {
                Titulo = "Control de cambios",
                Tipo = Seccion.TipoCambios,
                Ayuda = "Registre cada versión del documento.",
                FilasCambio = new ObservableCollection<FilaCambio>
                {
                    new()
                    {
                        Version = "01",
                        Fecha = "01/10/2026",
                        Descripcion = "Emisión inicial del documento alineado a la norma ISO 9001:2015.",
                        Elaboro = "Jefe de Logística",
                        Reviso = "Administrador",
                        Aprobo = "Gerente General"
                    }
                }
            }
        ];
    }

    private static DateTimeOffset Fecha(DateTime valor) =>
        new(DateTime.SpecifyKind(valor.Date, DateTimeKind.Local));
}
