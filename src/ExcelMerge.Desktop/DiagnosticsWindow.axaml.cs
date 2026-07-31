using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;

namespace ExcelMerge.Desktop;

public sealed partial class DiagnosticsWindow : Window
{
    public DiagnosticsWindow() => AvaloniaXamlLoader.Load(this);

    private void CloseWindow(object? sender, RoutedEventArgs e) => Close();
}
