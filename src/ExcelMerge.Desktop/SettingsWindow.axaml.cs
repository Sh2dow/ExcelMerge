using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using ExcelMerge.Application;

namespace ExcelMerge.Desktop;

public sealed partial class SettingsWindow : Window
{
    private readonly MainWindowViewModel? _viewModel;

    public SettingsWindow()
        : this(null)
    {
    }

    public SettingsWindow(MainWindowViewModel? viewModel)
    {
        _viewModel = viewModel;
        AvaloniaXamlLoader.Load(this);
        this.FindControl<ComboBox>("LanguageBox")!.SelectedIndex = LocalizationService.CurrentCulture switch
        {
            "zh-CN" => 1,
            _ => 0,
        };
        this.FindControl<ComboBox>("ThemeBox")!.SelectedIndex =
            (viewModel?.CurrentSettings.Theme ?? "System") switch
            {
                "Light" => 1,
                "Dark" => 2,
                _ => 0,
            };
        this.FindControl<NumericUpDown>("RowHeightBox")!.Value = (decimal)GridSettings.RowHeight;
        this.FindControl<NumericUpDown>("ColumnWidthBox")!.Value = (decimal)GridSettings.ColumnWidth;
        this.FindControl<NumericUpDown>("CacheRowsBox")!.Value = GridSettings.CacheRows;
        var settings = viewModel?.CurrentSettings ?? new ApplicationSettings();
        this.FindControl<TextBox>("KeyColumnsBox")!.Text = FormatKeyColumns(settings.KeyColumns);
        this.FindControl<CheckBox>("FormulaValuesBox")!.IsChecked = settings.CompareFormulaCachedValues;
        this.FindControl<CheckBox>("DisplayTextBox")!.IsChecked = settings.CompareDisplayText;
        this.FindControl<CheckBox>("CellStylesBox")!.IsChecked = settings.CompareCellStyles;
        this.FindControl<CheckBox>("RowMetadataBox")!.IsChecked = settings.CompareRowMetadata;
        this.FindControl<CheckBox>("BlankCellsBox")!.IsChecked = settings.TreatExplicitBlankCellsAsMissing;
        this.FindControl<NumericUpDown>("RowHeightBox")!.ValueChanged += GridValueChanged;
        this.FindControl<NumericUpDown>("ColumnWidthBox")!.ValueChanged += GridValueChanged;
        this.FindControl<NumericUpDown>("CacheRowsBox")!.ValueChanged += GridValueChanged;
    }

    private void LanguageChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox { SelectedItem: ComboBoxItem item } && item.Tag is string culture)
        {
            LocalizationService.Apply(culture);
        }
    }

    private void ThemeChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is not ComboBox { SelectedItem: ComboBoxItem item } || item.Tag is not string theme)
        {
            return;
        }

        App.ApplyTheme(theme);
        if (_viewModel is not null && _viewModel.CurrentSettings.Theme != theme)
        {
            _ = _viewModel.SaveSettingsAsync(_viewModel.CurrentSettings with { Theme = theme });
        }
    }

    private void GridValueChanged(object? sender, NumericUpDownValueChangedEventArgs e)
    {
        GridSettings.RowHeight = (double)(this.FindControl<NumericUpDown>("RowHeightBox")!.Value ?? 25);
        GridSettings.ColumnWidth = (double)(this.FindControl<NumericUpDown>("ColumnWidthBox")!.Value ?? 118);
        GridSettings.CacheRows = (int)(this.FindControl<NumericUpDown>("CacheRowsBox")!.Value ?? 256);
    }

    private async void SaveSettings(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is null)
        {
            Close();
            return;
        }

        try
        {
            var settings = _viewModel.CurrentSettings with
            {
                KeyColumns = ParseKeyColumns(this.FindControl<TextBox>("KeyColumnsBox")!.Text),
                CompareFormulaCachedValues = this.FindControl<CheckBox>("FormulaValuesBox")!.IsChecked == true,
                CompareDisplayText = this.FindControl<CheckBox>("DisplayTextBox")!.IsChecked == true,
                CompareCellStyles = this.FindControl<CheckBox>("CellStylesBox")!.IsChecked == true,
                CompareRowMetadata = this.FindControl<CheckBox>("RowMetadataBox")!.IsChecked == true,
                TreatExplicitBlankCellsAsMissing = this.FindControl<CheckBox>("BlankCellsBox")!.IsChecked == true,
            };
            await _viewModel.SaveSettingsAsync(settings);
            Close();
        }
        catch (FormatException)
        {
            this.FindControl<TextBlock>("ValidationText")!.Text = LocalizationService.Get("InvalidKeyColumns");
        }
        catch (Exception exception)
        {
            this.FindControl<TextBlock>("ValidationText")!.Text = exception.Message;
        }
    }

    private static int[] ParseKeyColumns(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return [];
        }

        var columns = new SortedSet<int>();
        foreach (var token in text.Split([',', ';', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries))
        {
            int column;
            if (int.TryParse(token, out var number) && number > 0)
            {
                column = number - 1;
            }
            else
            {
                column = 0;
                foreach (var character in token)
                {
                    if (!char.IsAsciiLetter(character))
                    {
                        throw new FormatException();
                    }

                    column = checked((column * 26) + char.ToUpperInvariant(character) - 'A' + 1);
                }

                column--;
            }

            if (column < 0 || !columns.Add(column))
            {
                throw new FormatException();
            }
        }

        return [.. columns];
    }

    private static string FormatKeyColumns(IReadOnlyList<int> columns) =>
        string.Join(", ", columns.Select(static column => ColumnName(column)));

    private static string ColumnName(int columnIndex)
    {
        Span<char> buffer = stackalloc char[8];
        var cursor = buffer.Length;
        var value = checked(columnIndex + 1);
        while (value > 0)
        {
            value--;
            buffer[--cursor] = (char)('A' + (value % 26));
            value /= 26;
        }

        return buffer[cursor..].ToString();
    }

    private void CloseWindow(object? sender, RoutedEventArgs e) => Close();
}
