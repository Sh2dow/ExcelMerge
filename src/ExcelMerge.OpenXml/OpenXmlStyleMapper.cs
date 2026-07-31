using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace ExcelMerge.OpenXml;

internal sealed class OpenXmlStyleMapper
{
    private readonly WorkbookPart _localWorkbookPart;
    private readonly WorkbookPart _remoteWorkbookPart;
    private readonly Dictionary<uint, uint> _cellFormatMap = [];
    private readonly Dictionary<uint, uint> _cellStyleFormatMap = [];
    private readonly Dictionary<uint, uint> _fontMap = [];
    private readonly Dictionary<uint, uint> _fillMap = [];
    private readonly Dictionary<uint, uint> _borderMap = [];
    private readonly Dictionary<uint, uint> _numberFormatMap = [];
    private readonly Dictionary<uint, uint> _differentialFormatMap = [];
    private bool? _themesEquivalent;

    public OpenXmlStyleMapper(WorkbookPart localWorkbookPart, WorkbookPart remoteWorkbookPart)
    {
        _localWorkbookPart = localWorkbookPart;
        _remoteWorkbookPart = remoteWorkbookPart;
    }

    public uint MapCellFormat(uint remoteIndex)
    {
        if (_cellFormatMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remoteStyles = _remoteWorkbookPart.WorkbookStylesPart?.Stylesheet;
        if (remoteStyles is null)
        {
            if (remoteIndex != 0)
            {
                throw InvalidStyles($"REMOTE style index {remoteIndex} has no style table.");
            }

            return 0;
        }

        var remoteFormat = GetAt<CellFormat>(
            remoteStyles.CellFormats,
            remoteIndex,
            "REMOTE cell format");
        var localStyles = EnsureLocalStyles();
        var clone = (CellFormat)remoteFormat.CloneNode(deep: true);
        MapCellFormatDependencies(clone, remoteStyles, mapBaseStyle: true);
        PrepareImportedStyle(clone);

        var existing = FindEquivalent(localStyles.CellFormats!, clone);
        mapped = existing ?? Append(localStyles.CellFormats!, clone);
        _cellFormatMap.Add(remoteIndex, mapped);
        SaveLocalStyles(localStyles);
        return mapped;
    }

    public uint MapDifferentialFormat(uint remoteIndex)
    {
        if (_differentialFormatMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remoteStyles = _remoteWorkbookPart.WorkbookStylesPart?.Stylesheet;
        var remoteFormat = GetAt<DifferentialFormat>(
            remoteStyles?.DifferentialFormats,
            remoteIndex,
            "REMOTE differential format");
        var localStyles = EnsureLocalStyles();
        localStyles.DifferentialFormats ??= new DifferentialFormats();
        var clone = (DifferentialFormat)remoteFormat.CloneNode(deep: true);
        if (clone.NumberingFormat is { } numberingFormat &&
            numberingFormat.NumberFormatId?.Value is { } numberFormatId)
        {
            numberingFormat.NumberFormatId = MapNumberFormat(numberFormatId, remoteStyles!);
        }

        PrepareImportedStyle(clone);
        var existing = FindEquivalent(localStyles.DifferentialFormats, clone);
        mapped = existing ?? Append(localStyles.DifferentialFormats, clone);
        _differentialFormatMap.Add(remoteIndex, mapped);
        SaveLocalStyles(localStyles);
        return mapped;
    }

    public void Save()
    {
        if (_localWorkbookPart.WorkbookStylesPart?.Stylesheet is { } stylesheet)
        {
            SaveLocalStyles(stylesheet);
        }
    }

    private void MapCellFormatDependencies(
        CellFormat format,
        Stylesheet remoteStyles,
        bool mapBaseStyle)
    {
        format.NumberFormatId = MapNumberFormat(format.NumberFormatId?.Value ?? 0, remoteStyles);
        format.FontId = MapFont(format.FontId?.Value ?? 0, remoteStyles);
        format.FillId = MapFill(format.FillId?.Value ?? 0, remoteStyles);
        format.BorderId = MapBorder(format.BorderId?.Value ?? 0, remoteStyles);
        if (mapBaseStyle && format.FormatId?.Value is { } baseStyleIndex)
        {
            format.FormatId = MapCellStyleFormat(baseStyleIndex, remoteStyles);
        }
    }

    private uint MapCellStyleFormat(uint remoteIndex, Stylesheet remoteStyles)
    {
        if (_cellStyleFormatMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remoteFormat = GetAt<CellFormat>(
            remoteStyles.CellStyleFormats,
            remoteIndex,
            "REMOTE cell style format");
        var localStyles = EnsureLocalStyles();
        var clone = (CellFormat)remoteFormat.CloneNode(deep: true);
        MapCellFormatDependencies(clone, remoteStyles, mapBaseStyle: false);
        PrepareImportedStyle(clone);
        var existing = FindEquivalent(localStyles.CellStyleFormats!, clone);
        mapped = existing ?? Append(localStyles.CellStyleFormats!, clone);
        _cellStyleFormatMap.Add(remoteIndex, mapped);
        return mapped;
    }

    private uint MapFont(uint remoteIndex, Stylesheet remoteStyles)
    {
        if (_fontMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remote = GetAt<Font>(remoteStyles.Fonts, remoteIndex, "REMOTE font");
        var localStyles = EnsureLocalStyles();
        var clone = (Font)remote.CloneNode(deep: true);
        PrepareImportedStyle(clone);
        var existing = FindEquivalent(localStyles.Fonts!, clone);
        mapped = existing ?? Append(localStyles.Fonts!, clone);
        _fontMap.Add(remoteIndex, mapped);
        return mapped;
    }

    private uint MapFill(uint remoteIndex, Stylesheet remoteStyles)
    {
        if (_fillMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remote = GetAt<Fill>(remoteStyles.Fills, remoteIndex, "REMOTE fill");
        var localStyles = EnsureLocalStyles();
        var clone = (Fill)remote.CloneNode(deep: true);
        PrepareImportedStyle(clone);
        var existing = FindEquivalent(localStyles.Fills!, clone);
        mapped = existing ?? Append(localStyles.Fills!, clone);
        _fillMap.Add(remoteIndex, mapped);
        return mapped;
    }

    private uint MapBorder(uint remoteIndex, Stylesheet remoteStyles)
    {
        if (_borderMap.TryGetValue(remoteIndex, out var mapped))
        {
            return mapped;
        }

        var remote = GetAt<Border>(remoteStyles.Borders, remoteIndex, "REMOTE border");
        var localStyles = EnsureLocalStyles();
        var clone = (Border)remote.CloneNode(deep: true);
        PrepareImportedStyle(clone);
        var existing = FindEquivalent(localStyles.Borders!, clone);
        mapped = existing ?? Append(localStyles.Borders!, clone);
        _borderMap.Add(remoteIndex, mapped);
        return mapped;
    }

    private uint MapNumberFormat(uint remoteId, Stylesheet remoteStyles)
    {
        if (_numberFormatMap.TryGetValue(remoteId, out var mapped))
        {
            return mapped;
        }

        var remoteFormat = remoteStyles.NumberingFormats?
            .Elements<NumberingFormat>()
            .FirstOrDefault(format => format.NumberFormatId?.Value == remoteId);
        if (remoteFormat is null)
        {
            _numberFormatMap.Add(remoteId, remoteId);
            return remoteId;
        }

        var localStyles = EnsureLocalStyles();
        localStyles.NumberingFormats ??= new NumberingFormats();
        var formatCode = remoteFormat.FormatCode?.Value ?? throw InvalidStyles(
            $"REMOTE number format {remoteId} has no format code.");
        var equivalent = localStyles.NumberingFormats.Elements<NumberingFormat>()
            .FirstOrDefault(format => string.Equals(
                format.FormatCode?.Value,
                formatCode,
                StringComparison.Ordinal));
        if (equivalent?.NumberFormatId?.Value is { } equivalentId)
        {
            mapped = equivalentId;
        }
        else
        {
            var usedIds = localStyles.NumberingFormats.Elements<NumberingFormat>()
                .Select(static format => format.NumberFormatId?.Value ?? 0U)
                .ToHashSet();
            mapped = 164;
            while (usedIds.Contains(mapped))
            {
                mapped = checked(mapped + 1);
            }

            var clone = (NumberingFormat)remoteFormat.CloneNode(deep: true);
            clone.NumberFormatId = mapped;
            localStyles.NumberingFormats.Append(clone);
            localStyles.NumberingFormats.Count = (uint)localStyles.NumberingFormats.ChildElements.Count;
        }

        _numberFormatMap.Add(remoteId, mapped);
        return mapped;
    }

    private Stylesheet EnsureLocalStyles()
    {
        var part = _localWorkbookPart.WorkbookStylesPart;
        if (part is null)
        {
            part = _localWorkbookPart.AddNewPart<WorkbookStylesPart>();
            part.Stylesheet = new Stylesheet();
        }

        var styles = part.Stylesheet;
        styles.Fonts ??= new Fonts(new Font());
        styles.Fills ??= new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        styles.Borders ??= new Borders(new Border());
        styles.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
        styles.CellFormats ??= new CellFormats(new CellFormat());
        styles.CellStyles ??= new CellStyles(new CellStyle
        {
            Name = "Normal",
            FormatId = 0,
            BuiltinId = 0,
        });
        return styles;
    }

    private void PrepareImportedStyle(OpenXmlElement element)
    {
        if (ThemesEquivalent())
        {
            return;
        }

        ResolveThemeColors(element);
        ResolveThemeFonts(element);
        if (UsesThemeReference(element))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "A REMOTE style contains a theme reference that cannot be resolved.");
        }
    }

    private void ResolveThemeColors(OpenXmlElement element)
    {
        foreach (var color in element.Descendants<ColorType>())
        {
            if (color.Theme?.Value is not { } themeIndex)
            {
                continue;
            }

            var rgb = GetThemeRgb(themeIndex);
            rgb = ApplyTint(rgb, color.Tint?.Value ?? 0);
            color.Theme = null;
            color.Tint = null;
            color.Rgb = "FF" + rgb;
        }
    }

    private void ResolveThemeFonts(OpenXmlElement element)
    {
        var fonts = element is Font rootFont
            ? new[] { rootFont }
            : element.Descendants<Font>();
        foreach (var font in fonts)
        {
            var scheme = font.FontScheme?.Val?.Value;
            if (scheme is null)
            {
                continue;
            }

            var themeFonts = _remoteWorkbookPart.ThemePart?
                .Theme?
                .ThemeElements?
                .FontScheme;
            var typeface = scheme == FontSchemeValues.Major
                ? themeFonts?.MajorFont?.LatinFont?.Typeface?.Value
                : scheme == FontSchemeValues.Minor
                    ? themeFonts?.MinorFont?.LatinFont?.Typeface?.Value
                    : null;
            if (string.IsNullOrWhiteSpace(typeface))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    $"The REMOTE theme does not define its '{scheme}' Latin font.");
            }

            font.FontName = new FontName { Val = typeface };
            font.FontScheme?.Remove();
        }
    }

    private string GetThemeRgb(uint themeIndex)
    {
        var scheme = _remoteWorkbookPart.ThemePart?
            .Theme?
            .ThemeElements?
            .ColorScheme ?? throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "The REMOTE style refers to a missing theme color scheme.");
        OpenXmlCompositeElement? entry = themeIndex switch
        {
            0 => scheme.Light1Color,
            1 => scheme.Dark1Color,
            2 => scheme.Light2Color,
            3 => scheme.Dark2Color,
            4 => scheme.Accent1Color,
            5 => scheme.Accent2Color,
            6 => scheme.Accent3Color,
            7 => scheme.Accent4Color,
            8 => scheme.Accent5Color,
            9 => scheme.Accent6Color,
            10 => scheme.Hyperlink,
            11 => scheme.FollowedHyperlinkColor,
            _ => null,
        };
        if (entry?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value is { Length: 6 } rgb)
        {
            return rgb.ToUpperInvariant();
        }

        if (entry?.GetFirstChild<Drawing.SystemColor>()?.LastColor?.Value is { Length: 6 } systemRgb)
        {
            return systemRgb.ToUpperInvariant();
        }

        throw new OpenXmlWriterException(
            OpenXmlWriterError.UnsupportedOperation,
            $"REMOTE theme color {themeIndex} is not represented as RGB or a system fallback color.");
    }

    private static string ApplyTint(string rgb, double tint)
    {
        if (!double.IsFinite(tint) || tint is < -1 or > 1)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                "A REMOTE theme color has an invalid tint.");
        }

        Span<char> result = stackalloc char[6];
        for (var index = 0; index < 3; index++)
        {
            var component = Convert.ToInt32(rgb.Substring(index * 2, 2), 16);
            var adjusted = tint < 0
                ? component * (1 + tint)
                : component + ((255 - component) * tint);
            var rounded = Math.Clamp((int)Math.Round(adjusted, MidpointRounding.AwayFromZero), 0, 255);
            rounded.TryFormat(result[(index * 2)..], out _, "X2");
        }

        return result.ToString();
    }

    private bool ThemesEquivalent()
    {
        if (_themesEquivalent.HasValue)
        {
            return _themesEquivalent.Value;
        }

        var local = _localWorkbookPart.ThemePart?.Theme?.OuterXml;
        var remote = _remoteWorkbookPart.ThemePart?.Theme?.OuterXml;
        _themesEquivalent = string.Equals(local, remote, StringComparison.Ordinal);
        return _themesEquivalent.Value;
    }

    private static bool UsesThemeReference(OpenXmlElement element)
    {
        if (element is FontScheme)
        {
            return true;
        }

        if (element.GetAttributes().Any(static attribute =>
            string.Equals(attribute.LocalName, "theme", StringComparison.Ordinal)))
        {
            return true;
        }

        foreach (var child in element.ChildElements)
        {
            if (UsesThemeReference(child))
            {
                return true;
            }
        }

        return false;
    }

    private static T GetAt<T>(OpenXmlCompositeElement? collection, uint index, string description)
        where T : OpenXmlElement
    {
        if (collection is null || index > int.MaxValue)
        {
            throw InvalidStyles($"{description} index {index} is missing.");
        }

        var value = collection.Elements<T>().ElementAtOrDefault((int)index);
        return value ?? throw InvalidStyles($"{description} index {index} is missing.");
    }

    private static uint? FindEquivalent<T>(OpenXmlCompositeElement collection, T candidate)
        where T : OpenXmlElement
    {
        var index = 0U;
        foreach (var value in collection.Elements<T>())
        {
            if (string.Equals(value.OuterXml, candidate.OuterXml, StringComparison.Ordinal))
            {
                return index;
            }

            index++;
        }

        return null;
    }

    private static uint Append<T>(OpenXmlCompositeElement collection, T value)
        where T : OpenXmlElement
    {
        var index = checked((uint)collection.Elements<T>().Count());
        collection.Append(value);
        SetCount(collection, index + 1);
        return index;
    }

    private static void SetCount(OpenXmlCompositeElement collection, uint count)
    {
        switch (collection)
        {
            case Fonts fonts:
                fonts.Count = count;
                break;
            case Fills fills:
                fills.Count = count;
                break;
            case Borders borders:
                borders.Count = count;
                break;
            case CellStyleFormats cellStyleFormats:
                cellStyleFormats.Count = count;
                break;
            case CellFormats cellFormats:
                cellFormats.Count = count;
                break;
            case DifferentialFormats differentialFormats:
                differentialFormats.Count = count;
                break;
        }
    }

    private static void SaveLocalStyles(Stylesheet styles)
    {
        if (styles.Fonts is { } fonts)
            fonts.Count = (uint)fonts.ChildElements.Count;
        if (styles.Fills is { } fills)
            fills.Count = (uint)fills.ChildElements.Count;
        if (styles.Borders is { } borders)
            borders.Count = (uint)borders.ChildElements.Count;
        if (styles.CellStyleFormats is { } cellStyleFormats)
            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;
        if (styles.CellFormats is { } cellFormats)
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
        if (styles.CellStyles is { } cellStyles)
            cellStyles.Count = (uint)cellStyles.ChildElements.Count;
        if (styles.DifferentialFormats is { } differentialFormats)
            differentialFormats.Count = (uint)differentialFormats.ChildElements.Count;
        styles.Save();
    }

    private static OpenXmlWriterException InvalidStyles(string message) =>
        new(OpenXmlWriterError.InvalidPackage, message);
}
