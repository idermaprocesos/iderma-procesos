using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Iderma_Procesos.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Iderma_Procesos.Services;

public static class WordDocumentService
{
    private const int Ancho = 10488;
    private const int ColLogo = 2103;
    private const int ColTitulo = 5230;
    private const int ColEtiqueta = 1400;
    private const int ColValor = 1755;
    private const int ColFirma = 3496;

    private static readonly CultureInfo Cultura = CultureInfo.GetCultureInfo("es-PE");

    public static void Crear(Stream destino, Procedimiento procedimiento, byte[] logoPng)
        => CrearCombinado(destino, [procedimiento], logoPng);

    public static void CrearCombinado(Stream destino, IReadOnlyList<Procedimiento> procedimientos, byte[] logoPng)
    {
        if (procedimientos.Count == 0)
            throw new InvalidOperationException("Seleccione al menos un procedimiento.");

        var temporal = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            CrearArchivo(temporal, procedimientos, logoPng);
            using var origen = File.OpenRead(temporal);
            origen.CopyTo(destino);
        }
        finally
        {
            if (File.Exists(temporal))
                File.Delete(temporal);
        }
    }

    public static void ExportarPdf(string docx, string pdf, Action? alCrear = null)
    {
        var tipo = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("No está instalado Microsoft Word, que se usa para crear el PDF.");

        dynamic word = Activator.CreateInstance(tipo)
            ?? throw new InvalidOperationException("No se pudo abrir Microsoft Word.");
        dynamic? documento = null;
        try
        {
            word.Visible = false;
            word.DisplayAlerts = 0;
            documento = word.Documents.Open(docx);
            int secciones = documento.Sections.Count;
            for (var s = 1; s <= secciones; s++)
            {
                dynamic campos = documento.Sections[s].Headers[1].Range.Fields;
                int total = campos.Count;
                for (var i = 1; i <= total; i++)
                    campos[i].Update();
            }

            documento.ExportAsFixedFormat(pdf, 17);
            documento.Close(false);
            documento = null;
            alCrear?.Invoke();
        }
        finally
        {
            if (documento is not null)
            {
                try { documento.Close(false); } catch { /* El PDF ya quedó creado. */ }
            }

            try { word.Quit(); } catch { /* Word puede tardar en cerrarse. */ }
        }
    }

    private static void CrearArchivo(string ruta, IReadOnlyList<Procedimiento> procedimientos, byte[] logoPng)
    {
        var (pxAncho, pxAlto) = TamanoPng(logoPng);

        using var documento = WordprocessingDocument.Create(ruta, WordprocessingDocumentType.Document);
        var principal = documento.AddMainDocumentPart();
        principal.Document = new Document(new Body());

        var primero = procedimientos[0];
        documento.PackageProperties.Title = procedimientos.Count == 1 ? primero.Titulo : "Procedimientos combinados";
        documento.PackageProperties.Subject = string.Join(", ", procedimientos.Select(p => p.Codigo).Where(c => !string.IsNullOrWhiteSpace(c)));
        documento.PackageProperties.Creator = "Iderma Capilar";
        documento.PackageProperties.LastModifiedBy = "Iderma Capilar";

        var cuerpo = principal.Document.Body!;
        var idDibujo = 1u;
        for (var indice = 0; indice < procedimientos.Count; indice++)
        {
            var procedimiento = procedimientos[indice];
            var idEncabezado = AgregarEncabezado(principal, procedimiento, logoPng, pxAncho, pxAlto, ref idDibujo);
            AgregarCuerpo(cuerpo, principal, procedimiento, logoPng, pxAncho, pxAlto, ref idDibujo);

            var cierre = indice < procedimientos.Count - 1;
            var propiedades = PropiedadesSeccion(idEncabezado, cierre);
            if (cierre)
                cuerpo.Append(new Paragraph(new ParagraphProperties(propiedades)));
            else
                cuerpo.Append(propiedades);
        }

        var estilos = principal.AddNewPart<StyleDefinitionsPart>();
        estilos.Styles = Estilos();
        estilos.Styles.Save();

        var ajustes = principal.AddNewPart<DocumentSettingsPart>();
        ajustes.Settings = new Settings(
            new DoNotDisplayPageBoundaries(),
            new UpdateFieldsOnOpen { Val = true },
            new Compatibility(
                new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "15"
                }));
        ajustes.Settings.Save();
        principal.Document.Save();
        documento.Dispose();
        CorregirTiposDeContenido(ruta);
    }

    private static void CorregirTiposDeContenido(string ruta)
    {
        const string tipoDocumento = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        var temporal = ruta + ".tmp";
        using (var origen = ZipFile.OpenRead(ruta))
        using (var destino = ZipFile.Open(temporal, ZipArchiveMode.Create))
        {
            foreach (var entrada in origen.Entries)
            {
                using var lectura = entrada.Open();
                using var salida = destino.CreateEntry(entrada.FullName, CompressionLevel.Optimal).Open();
                if (entrada.FullName != "[Content_Types].xml")
                {
                    lectura.CopyTo(salida);
                    continue;
                }

                using var lector = new StreamReader(lectura);
                var xml = lector.ReadToEnd();
                xml = xml.Replace(
                    $"<Default Extension=\"xml\" ContentType=\"{tipoDocumento}\" />",
                    "<Default Extension=\"xml\" ContentType=\"application/xml\" />");
                if (!xml.Contains("PartName=\"/word/document.xml\""))
                {
                    xml = xml.Replace(
                        "</Types>",
                        $"<Override PartName=\"/word/document.xml\" ContentType=\"{tipoDocumento}\" /></Types>");
                }

                using var escritor = new StreamWriter(salida);
                escritor.Write(xml);
            }
        }

        File.Delete(ruta);
        File.Move(temporal, ruta);
    }

    private static string AgregarEncabezado(
        MainDocumentPart principal,
        Procedimiento procedimiento,
        byte[] logoPng,
        int pxAncho,
        int pxAlto,
        ref uint idDibujo)
    {
        var encabezado = principal.AddNewPart<HeaderPart>();
        var idLogo = AgregarImagen(encabezado, logoPng);
        encabezado.Header = new Header(
            TablaControl(procedimiento, idLogo, pxAncho, pxAlto, idDibujo++),
            ParrafoMinimo());
        encabezado.Header.Save();
        return principal.GetIdOfPart(encabezado);
    }

    private static void AgregarCuerpo(
        Body cuerpo,
        MainDocumentPart principal,
        Procedimiento procedimiento,
        byte[] logoPng,
        int pxAncho,
        int pxAlto,
        ref uint idDibujo)
    {
        var idLogoCuerpo = AgregarImagen(principal, logoPng);
        var (cxLogo, cyLogo) = Medida(pxAncho, pxAlto, 4.6);
        cuerpo.Append(Espaciador(1500));
        cuerpo.Append(ParrafoImagen(idLogoCuerpo, cxLogo, cyLogo, idDibujo++, "Logo Iderma Capilar"));
        cuerpo.Append(Parrafo(procedimiento.TituloMayusculas, true, 22, JustificationValues.Center, antes: 160));
        cuerpo.Append(Espaciador(2400));
        cuerpo.Append(TablaFirmas(procedimiento));
        cuerpo.Append(ParrafoMinimo());

        var secciones = procedimiento.Secciones
            .Where(s => !string.IsNullOrWhiteSpace(s.Titulo)
                || !string.IsNullOrWhiteSpace(s.Contenido)
                || s.FilasDesarrollo is { Count: > 0 }
                || s.FilasCambio is { Count: > 0 })
            .ToList();

        for (var i = 0; i < secciones.Count; i++)
            cuerpo.Append(BloqueSeccion(secciones[i], i + 1, saltoDePagina: i == 0));
    }

    private static Table TablaControl(Procedimiento procedimiento, string idImagen, int pxAncho, int pxAlto, uint idDibujo)
    {
        var (cx, cy) = Medida(pxAncho, pxAlto, 1.15);
        var tabla = NuevaTabla(Ancho, conBordes: true, margenVertical: "0");
        tabla.Append(new TableGrid(
            new GridColumn { Width = ColLogo.ToString() },
            new GridColumn { Width = ColTitulo.ToString() },
            new GridColumn { Width = ColEtiqueta.ToString() },
            new GridColumn { Width = ColValor.ToString() }));

        var filas = new[]
        {
            ("Código:", procedimiento.Codigo, true),
            ("Versión:", procedimiento.Version, false),
            ("Fecha:", procedimiento.FechaDocumento.ToString("dd/MM/yyyy", Cultura), false),
            ("Página:", "", false)
        };

        for (var i = 0; i < filas.Length; i++)
        {
            var fila = new TableRow();
            fila.Append(new TableRowProperties(
                new TableRowHeight { Val = 280, HeightType = HeightRuleValues.Exact },
                new CantSplit()));

            fila.Append(CeldaFusion(ColLogo, i, i == 0
                ? ParrafoImagen(idImagen, cx, cy, idDibujo, "Logo Iderma Capilar", ajustado: true)
                : ParrafoMinimo()));

            fila.Append(CeldaFusion(ColTitulo, i, i == 0
                ? Parrafo(procedimiento.TituloMayusculas, true, 22, JustificationValues.Center, ajustado: true)
                : ParrafoMinimo()));

            fila.Append(Celda(ColEtiqueta, Parrafo(filas[i].Item1, true, 22, JustificationValues.Right, ajustado: true), completa: true));

            var valor = i == 3
                ? ParrafoPagina()
                : Parrafo(filas[i].Item2, false, 22, JustificationValues.Left, ajustado: true);
            fila.Append(Celda(ColValor, valor, completa: true));
            tabla.Append(fila);
        }

        return tabla;
    }

    private static Table TablaFirmas(Procedimiento procedimiento)
    {
        var tabla = NuevaTabla(Ancho, conBordes: true, margenVertical: "20");
        tabla.Append(new TableGrid(
            new GridColumn { Width = ColFirma.ToString() },
            new GridColumn { Width = ColFirma.ToString() },
            new GridColumn { Width = ColFirma.ToString() }));

        var vacia = new TableRow();
        vacia.Append(new TableRowProperties(
            new TableRowHeight { Val = 2000, HeightType = HeightRuleValues.Exact },
            new CantSplit()));
        vacia.Append(Celda(ColFirma, ParrafoVacio(), completa: true));
        vacia.Append(Celda(ColFirma, ParrafoVacio(), completa: true));
        vacia.Append(Celda(ColFirma, ParrafoVacio(), completa: true));
        tabla.Append(vacia);

        var fechas = new TableRow();
        fechas.Append(new TableRowProperties(
            new TableRowHeight { Val = 320, HeightType = HeightRuleValues.Exact },
            new CantSplit()));
        fechas.Append(Celda(ColFirma, Parrafo("Fecha: " + procedimiento.FechaElaboracionTexto, true, 22, JustificationValues.Center, ajustado: true), completa: true));
        fechas.Append(Celda(ColFirma, Parrafo("Fecha: " + procedimiento.FechaRevisionTexto, true, 22, JustificationValues.Center, ajustado: true), completa: true));
        fechas.Append(Celda(ColFirma, Parrafo("Fecha: " + procedimiento.FechaAprobacionTexto, true, 22, JustificationValues.Center, ajustado: true), completa: true));
        tabla.Append(fechas);

        var datos = new TableRow();
        datos.Append(new TableRowProperties(
            new TableRowHeight { Val = 960, HeightType = HeightRuleValues.Exact },
            new CantSplit()));
        datos.Append(CeldaFirma(procedimiento.ElaboradoNombre, procedimiento.ElaboradoCargo, "Elaborado por:"));
        datos.Append(CeldaFirma(procedimiento.RevisadoNombre, procedimiento.RevisadoCargo, "Revisado por:"));
        datos.Append(CeldaFirma(procedimiento.AprobadoNombre, procedimiento.AprobadoCargo, "Aprobado por:"));
        tabla.Append(datos);
        return tabla;
    }

    private static TableCell CeldaFirma(string nombre, string cargo, string etiqueta)
    {
        return Celda(ColFirma, null, completa: true,
            Parrafo(etiqueta, true, 22, JustificationValues.Center, ajustado: true),
            Parrafo(nombre, false, 22, JustificationValues.Center, ajustado: true),
            Parrafo(cargo, false, 22, JustificationValues.Center, ajustado: true));
    }

    private static IEnumerable<OpenXmlElement> BloqueSeccion(Seccion seccion, int numero, bool saltoDePagina)
    {
        var titulo = Regex.Replace((seccion.Titulo ?? "").Trim(), @"^\d+\s*[\.\-]\s*", "");
        if (string.IsNullOrWhiteSpace(titulo))
            titulo = "Procedimiento";

        yield return Parrafo(
            $"{numero}. {titulo}".ToUpper(Cultura),
            true,
            22,
            JustificationValues.Left,
            antes: saltoDePagina ? 80 : 280,
            despues: 80,
            interlineado: "360",
            salto: saltoDePagina);

        foreach (var parrafo in ParrafosCuerpo(seccion.Contenido))
            yield return parrafo;

        if (seccion.Tipo == Seccion.TipoDesarrollo)
        {
            seccion.AsegurarFilas();
            var filas = seccion.FilasDesarrollo
                .Select(f => new[] { f.Area, f.Responsable, f.Actividad })
                .ToList();
            yield return TablaSimple(
                [2200, 2600, 5688],
                ["Área", "Responsable", "Actividad Detallada"],
                filas,
                [false, false, true]);
        }
        else if (seccion.Tipo == Seccion.TipoCambios)
        {
            seccion.AsegurarFilas();
            var filas = seccion.FilasCambio
                .Select(f => new[] { f.Version, f.Fecha, f.Descripcion, f.Elaboro, f.Reviso, f.Aprobo })
                .ToList();
            yield return TablaSimple(
                [980, 1600, 3444, 1488, 1488, 1488],
                ["Versión", "Fecha", "Descripción del cambio", "Elaboró", "Revisó", "Aprobó"],
                filas,
                [false, false, true, false, false, false]);
        }
    }

    private static IEnumerable<Paragraph> ParrafosCuerpo(string? contenido)
    {
        var lineas = (contenido ?? "").Replace("\r\n", "\n").Split('\n');
        var escribio = false;
        foreach (var linea in lineas)
        {
            if (string.IsNullOrWhiteSpace(linea))
            {
                if (escribio)
                    yield return Parrafo("", false, 22, JustificationValues.Both, interlineado: "360");
                continue;
            }

            escribio = true;
            yield return Parrafo(linea.TrimEnd(), false, 22, JustificationValues.Both, despues: 60, interlineado: "360");
        }
    }

    private static Table TablaSimple(int[] anchos, string[] encabezados, List<string[]> filas, bool[] justificar)
    {
        var tabla = NuevaTabla(Ancho, conBordes: true, margenVertical: "40");
        var grilla = new TableGrid();
        foreach (var ancho in anchos)
            grilla.Append(new GridColumn { Width = ancho.ToString() });
        tabla.Append(grilla);

        var encabezado = new TableRow();
        encabezado.Append(new TableRowProperties(new TableHeader()));
        for (var i = 0; i < encabezados.Length; i++)
            encabezado.Append(CeldaTexto(anchos[i], encabezados[i], true, JustificationValues.Center));
        tabla.Append(encabezado);

        if (filas.Count == 0)
            filas.Add(Enumerable.Repeat("", anchos.Length).ToArray());

        foreach (var fila in filas)
        {
            var tr = new TableRow();
            for (var i = 0; i < anchos.Length; i++)
            {
                var texto = i < fila.Length ? fila[i] : "";
                var alineacion = justificar[i] ? JustificationValues.Both : JustificationValues.Center;
                tr.Append(CeldaTexto(anchos[i], texto ?? "", false, alineacion));
            }

            tabla.Append(tr);
        }

        return tabla;
    }

    private static TableCell CeldaTexto(int ancho, string texto, bool negrita, JustificationValues alineacion)
    {
        var lineas = (texto ?? "").Replace("\r\n", "\n").Split('\n');
        var parrafos = lineas
            .Select(linea => Parrafo(linea.TrimEnd(), negrita, 22, alineacion, antes: 40, despues: 40, interlineado: "360"))
            .ToArray();
        if (parrafos.Length == 0)
            parrafos = [Parrafo("", negrita, 22, alineacion, antes: 40, despues: 40, interlineado: "360")];

        return Celda(ancho, parrafos[0], true, parrafos.Skip(1).ToArray());
    }

    private static SectionProperties PropiedadesSeccion(string idEncabezado, bool saltoDeSeccion)
    {
        var propiedades = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = idEncabezado });

        if (saltoDeSeccion)
            propiedades.Append(new SectionType { Val = SectionMarkValues.NextPage });

        propiedades.Append(new PageSize { Width = 11906, Height = 16838, Orient = PageOrientationValues.Portrait });
        propiedades.Append(new PageMargin
        {
            Top = 2280,
            Bottom = 700,
            Left = 709,
            Right = 709,
            Header = 709,
            Footer = 284,
            Gutter = 0
        });
        propiedades.Append(new PageNumberType { Start = 1 });
        propiedades.Append(new TitlePage { Val = false });
        return propiedades;
    }

    private static Styles Estilos()
    {
        return new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" },
                        new FontSize { Val = "22" },
                        new FontSizeComplexScript { Val = "22" },
                        new Languages { Val = "es-PE" })),
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }))));
    }

    private static Table NuevaTabla(int ancho, bool conBordes, string margenVertical = "40")
    {
        var bordes = conBordes
            ? Bordes(BorderValues.Single)
            : Bordes(BorderValues.None);

        var tabla = new Table();
        tabla.Append(new TableProperties(
            new TableWidth { Width = ancho.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableJustification { Val = TableRowAlignmentValues.Center },
            new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa },
            bordes,
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableCellMarginDefault(
                new TopMargin { Width = margenVertical, Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "60", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = margenVertical, Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "60", Type = TableWidthUnitValues.Dxa })));
        return tabla;
    }

    private static TableBorders Bordes(BorderValues valor)
    {
        return new TableBorders(
            new TopBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
            new LeftBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
            new BottomBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
            new RightBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
            new InsideHorizontalBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
            new InsideVerticalBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 });
    }

    private static TableCell Celda(int ancho, Paragraph? parrafo, bool completa, params Paragraph[] extra)
    {
        var celda = new TableCell();
        celda.Append(new TableCellProperties(
            new TableCellWidth { Width = ancho.ToString(), Type = TableWidthUnitValues.Dxa },
            BordesCelda(BorderValues.Single),
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }));

        if (parrafo is not null)
            celda.Append(parrafo);
        foreach (var item in extra)
            celda.Append(item);
        if (parrafo is null && extra.Length == 0)
            celda.Append(ParrafoVacio());
        return celda;
    }

    private static TableCell CeldaFusion(int ancho, int fila, Paragraph contenido)
    {
        var arriba = fila == 0 ? BorderValues.Single : BorderValues.Nil;
        var abajo = fila == 3 ? BorderValues.Single : BorderValues.Nil;
        var celda = new TableCell();
        var merge = new VerticalMerge();
        if (fila == 0)
            merge.Val = MergedCellValues.Restart;

        celda.Append(new TableCellProperties(
            new TableCellWidth { Width = ancho.ToString(), Type = TableWidthUnitValues.Dxa },
            merge,
            new TableCellBorders(
                new TopBorder { Val = arriba, Size = 4U, Color = "000000", Space = 0 },
                new LeftBorder { Val = BorderValues.Single, Size = 4U, Color = "000000", Space = 0 },
                new BottomBorder { Val = abajo, Size = 4U, Color = "000000", Space = 0 },
                new RightBorder { Val = BorderValues.Single, Size = 4U, Color = "000000", Space = 0 }),
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }));
        celda.Append(contenido);
        return celda;
    }

    private static TableCellBorders BordesCelda(BorderValues valor) => new(
        new TopBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
        new LeftBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
        new BottomBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 },
        new RightBorder { Val = valor, Size = 4U, Color = "000000", Space = 0 });

    private static Paragraph Parrafo(
        string texto,
        bool negrita,
        int mediosPuntos,
        JustificationValues alineacion,
        int antes = 0,
        int despues = 0,
        string interlineado = "240",
        bool salto = false,
        bool ajustado = false)
    {
        var propiedades = new ParagraphProperties();
        if (salto)
            propiedades.Append(new PageBreakBefore());
        propiedades.Append(new SpacingBetweenLines
        {
            Before = ajustado ? "0" : antes.ToString(),
            After = ajustado ? "0" : despues.ToString(),
            Line = ajustado ? (mediosPuntos * 10).ToString() : interlineado,
            LineRule = ajustado ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.Auto
        });
        propiedades.Append(new Justification { Val = alineacion });

        return new Paragraph(propiedades, new Run(Fuente(negrita, mediosPuntos), new Text(texto) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph ParrafoVacio()
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact }),
            new Run(new RunProperties(new FontSize { Val = "2" })));
    }

    private static Paragraph ParrafoMinimo()
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0", Line = "1", LineRule = LineSpacingRuleValues.Exact },
                new Justification { Val = JustificationValues.Left }),
            new Run(new RunProperties(
                new FontSize { Val = "2" },
                new FontSizeComplexScript { Val = "2" })));
    }

    private static Paragraph Espaciador(int twips)
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0", Line = twips.ToString(), LineRule = LineSpacingRuleValues.Exact }),
            new Run(new RunProperties(new FontSize { Val = "2" })));
    }

    private static Paragraph ParrafoPagina()
    {
        var parrafo = new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0", Line = "220", LineRule = LineSpacingRuleValues.Exact },
                new Justification { Val = JustificationValues.Left }));

        void Campo(string instruccion, string cache)
        {
            parrafo.Append(new Run(Fuente(false, 22), new FieldChar { FieldCharType = FieldCharValues.Begin }));
            parrafo.Append(new Run(Fuente(false, 22), new FieldCode(" " + instruccion + " ") { Space = SpaceProcessingModeValues.Preserve }));
            parrafo.Append(new Run(Fuente(false, 22), new FieldChar { FieldCharType = FieldCharValues.Separate }));
            parrafo.Append(new Run(Fuente(false, 22), new Text(cache)));
            parrafo.Append(new Run(Fuente(false, 22), new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        Campo("PAGE", "1");
        parrafo.Append(new Run(Fuente(false, 22), new Text(" de ") { Space = SpaceProcessingModeValues.Preserve }));
        Campo("SECTIONPAGES", "1");
        return parrafo;
    }

    private static Paragraph ParrafoImagen(string relacion, long cx, long cy, uint id, string nombre, bool ajustado = false)
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                new Justification { Val = JustificationValues.Center }),
            new Run(Fuente(false, ajustado ? 2 : 22), Imagen(relacion, cx, cy, id, nombre)));
    }

    private static RunProperties Fuente(bool negrita, int mediosPuntos)
    {
        var propiedades = new RunProperties(
            new RunFonts { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" });
        if (negrita)
        {
            propiedades.Append(new Bold());
            propiedades.Append(new BoldComplexScript());
        }

        propiedades.Append(new Color { Val = "000000" });
        propiedades.Append(new FontSize { Val = mediosPuntos.ToString() });
        propiedades.Append(new FontSizeComplexScript { Val = mediosPuntos.ToString() });
        propiedades.Append(new Languages { Val = "es-PE" });
        return propiedades;
    }

    private static Drawing Imagen(string relacion, long cx, long cy, uint id, string nombre)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = id, Name = nombre },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = id, Name = nombre + ".png" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relacion, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0,
                DistanceFromBottom = 0,
                DistanceFromLeft = 0,
                DistanceFromRight = 0
            });
    }

    private static string AgregarImagen(OpenXmlPartContainer contenedor, byte[] png)
    {
        ImagePart parte = contenedor switch
        {
            HeaderPart encabezado => encabezado.AddImagePart(ImagePartType.Png),
            MainDocumentPart principal => principal.AddImagePart(ImagePartType.Png),
            _ => throw new InvalidOperationException("No se puede insertar el logo.")
        };

        using var origen = new MemoryStream(png);
        parte.FeedData(origen);
        return contenedor.GetIdOfPart(parte);
    }

    private static (long Cx, long Cy) Medida(int pxAncho, int pxAlto, double anchoCm)
    {
        var cx = (long)Math.Round(anchoCm * 360000d);
        var cy = (long)Math.Round(cx * pxAlto / (double)pxAncho);
        return (cx, cy);
    }

    private static (int Ancho, int Alto) TamanoPng(byte[] png)
    {
        if (png.Length < 24 || png[0] != 0x89)
            return (1240, 928);

        var ancho = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
        var alto = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
        return (ancho, alto);
    }
}
