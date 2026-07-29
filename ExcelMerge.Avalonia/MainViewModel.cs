using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelMerge.Avalonia;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private string _basePath = string.Empty;
    private string _localPath = string.Empty;
    private string _remotePath = string.Empty;
    public string BasePath { get => _basePath; set => Set(ref _basePath, value); }
    public string LocalPath { get => _localPath; set => Set(ref _localPath, value); }
    public string RemotePath { get => _remotePath; set => Set(ref _remotePath, value); }
    public ApplicationMode Mode { get; }
    public bool IsMergeMode => Mode == ApplicationMode.Merge;
    public string ModeLabel => IsMergeMode ? "MERGE PREVIEW" : "DIFF";
    public RelayCommand CompareCommand { get; }
    public MainViewModel(Func<Task> compare, ApplicationMode mode)
    {
        Mode = mode;
        CompareCommand = new RelayCommand(() => _ = compare());
    }
    public event PropertyChangedEventHandler? PropertyChanged;
    private void Set(ref string field, string value, [CallerMemberName] string? name = null)
    {
        if (field == value) return;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}

public sealed class RelayCommand : System.Windows.Input.ICommand
{
    private readonly Action _execute;
    public RelayCommand(Action execute) => _execute = execute;
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => _execute();
    public event EventHandler? CanExecuteChanged { add { } remove { } }
}
