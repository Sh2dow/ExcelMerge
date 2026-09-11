using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Data.Converters;

namespace ExcelMerge.Desktop.Converters;

public sealed class BoolToGridLengthConverter : IValueConverter
{
    public GridLength TrueValue { get; set; } = new GridLength(1, GridUnitType.Star);

    public GridLength FalseValue { get; set; } = new GridLength(0);

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is true ? TrueValue : FalseValue;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
