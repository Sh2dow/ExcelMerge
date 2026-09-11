using Avalonia.Media;
using Avalonia.Styling;

namespace ExcelMerge.Desktop;

/// <summary>Theme-aware brushes and pens for the custom-drawn grid and change map.</summary>
public sealed class GridPalette
{
    private GridPalette(bool dark)
    {
        Background = Brush(dark ? "#1E2023" : "#FFFFFF");
        HeaderBackground = Brush(dark ? "#26292D" : "#E9ECEF");
        GridLine = Brush(dark ? "#3A3F45" : "#D5D9DD");
        Text = Brush(dark ? "#E6E8EB" : "#202428");
        MutedText = Brush(dark ? "#9AA3AD" : "#68717B");
        LocalHeader = Brush(dark ? "#1F3A2E" : "#DDEFE5");
        RemoteHeader = Brush(dark ? "#3D2430" : "#F2DFE4");
        ChangedBackground = Brush(dark ? "#4A3F1E" : "#FFF4C7");
        ConflictBackground = Brush(dark ? "#4A2430" : "#FADCE2");
        ChangedHighlight = Brush(dark ? "#8A6D1F" : "#F5D76E");
        ConflictHighlight = Brush(dark ? "#96445C" : "#F0A4B8");
        ResolvedBackground = Brush(dark ? "#1F3A2E" : "#DDEFE5");
        SelectedBackground = Brush(dark ? "#1F3A55" : "#DCEBFA");
        GridPen = Pen(GridLine, 1);
        SplitPen = Pen(Brush(dark ? "#4A525B" : "#8A939D"), 1.5);
        SelectionPen = Pen(Brush(dark ? "#4A96CC" : "#1769AA"), 2);
        ConflictPen = Pen(Brush(dark ? "#C96A82" : "#C54A68"), 1.5);
        ResolvedPen = Pen(Brush(dark ? "#5FBF8F" : "#4F8B68"), 1.5);
        MapBackground = Brush(dark ? "#232629" : "#EEF0F1");
        MapChanged = Brush(dark ? "#8A6D1F" : "#E4B928");
        MapConflict = Brush(dark ? "#96445C" : "#C54A68");
        MapResolved = Brush(dark ? "#2E6B4E" : "#4F8B68");
        MapViewportFill = Brush(dark ? "#35000000" : "#35FFFFFF");
        MapBorderPen = Pen(Brush(dark ? "#4A525B" : "#B8BEC4"), 1);
        MapViewportPen = Pen(Brush(dark ? "#C8CFD6" : "#30363C"), 1.5);
    }

    public static GridPalette Light { get; } = new(dark: false);

    public static GridPalette Dark { get; } = new(dark: true);

    public IBrush Background { get; }
    public IBrush HeaderBackground { get; }
    public IBrush GridLine { get; }
    public IBrush Text { get; }
    public IBrush MutedText { get; }
    public IBrush LocalHeader { get; }
    public IBrush RemoteHeader { get; }
    public IBrush ChangedBackground { get; }
    public IBrush ConflictBackground { get; }
    public IBrush ChangedHighlight { get; }
    public IBrush ConflictHighlight { get; }
    public IBrush ResolvedBackground { get; }
    public IBrush SelectedBackground { get; }
    public Pen GridPen { get; }
    public Pen SplitPen { get; }
    public Pen SelectionPen { get; }
    public Pen ConflictPen { get; }
    public Pen ResolvedPen { get; }
    public IBrush MapBackground { get; }
    public IBrush MapChanged { get; }
    public IBrush MapConflict { get; }
    public IBrush MapResolved { get; }
    public IBrush MapViewportFill { get; }
    public Pen MapBorderPen { get; }
    public Pen MapViewportPen { get; }

    public static GridPalette ForVariant(ThemeVariant? variant) =>
        variant == ThemeVariant.Dark ? Dark : Light;

    private static IBrush Brush(string color) => new SolidColorBrush(Color.Parse(color));

    private static Pen Pen(IBrush brush, double thickness) => new(brush, thickness);
}
