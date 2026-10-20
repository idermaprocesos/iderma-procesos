using Iderma_Procesos.Models;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace Iderma_Procesos;

public sealed class SelectorPlantillaSeccion : DataTemplateSelector
{
    public DataTemplate? Texto { get; set; }

    public DataTemplate? Desarrollo { get; set; }

    public DataTemplate? Cambios { get; set; }

    protected override DataTemplate? SelectTemplateCore(object item) => Elegir(item);

    protected override DataTemplate? SelectTemplateCore(object item, DependencyObject container) => Elegir(item);

    private DataTemplate? Elegir(object item)
    {
        if (item is not Seccion seccion)
            return Texto;

        return seccion.Tipo switch
        {
            Seccion.TipoDesarrollo => Desarrollo,
            Seccion.TipoCambios => Cambios,
            _ => Texto
        };
    }
}
