using Avalonia;

namespace ExcelMerge.Desktop;

internal static class LocalizationService
{
    private static readonly IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> Cultures =
        new Dictionary<string, IReadOnlyDictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["en-US"] = new Dictionary<string, string>
            {
                ["AppTitle"] = "ExcelMerge",
                ["Compare"] = "Compare",
                ["Merge"] = "Merge",
                ["Open"] = "Run",
                ["Cancel"] = "Cancel",
                ["SaveResult"] = "Save result",
                ["Settings"] = "Settings",
                ["Diagnostics"] = "Diagnostics",
                ["Base"] = "BASE",
                ["Local"] = "LOCAL",
                ["Remote"] = "REMOTE",
                ["Browse"] = "Browse",
                ["Sheets"] = "Worksheets",
                ["Recent"] = "Recent",
                ["Search"] = "Search",
                ["Exact"] = "Exact",
                ["CaseSensitive"] = "Case",
                ["Regex"] = "Regex",
                ["Previous"] = "Previous",
                ["Next"] = "Next",
                ["Changes"] = "Changes",
                ["Conflicts"] = "Conflicts",
                ["Inspector"] = "Inspector",
                ["Result"] = "RESULT",
                ["Formula"] = "Formula",
                ["Type"] = "Type",
                ["Value"] = "Value",
                ["UseLocal"] = "Use LOCAL",
                ["UseRemote"] = "Use REMOTE",
                ["UseBoth"] = "Use BOTH",
                ["UseCustom"] = "Use custom",
                ["CustomValue"] = "Custom value",
                ["CopyTsv"] = "Copy TSV",
                ["CopyCsv"] = "Copy CSV",
                ["Ready"] = "Ready",
                ["NoSession"] = "Choose files and run a comparison.",
                ["Language"] = "Language",
                ["Theme"] = "Theme",
                ["Light"] = "Light",
                ["Dark"] = "Dark",
                ["System"] = "System",
                ["Grid"] = "Grid",
                ["RowHeight"] = "Row height",
                ["ColumnWidth"] = "Column width",
                ["CacheRows"] = "Cached rows",
                ["Close"] = "Close",
                ["OperationLog"] = "Operation log",
                ["Unresolved"] = "Unresolved",
                ["Resolved"] = "Resolved",
                ["Loading"] = "Loading workbooks",
                ["Saving"] = "Saving result",
                ["Completed"] = "Completed",
                ["Failed"] = "Operation failed",
                ["Swap"] = "Swap sides",
                ["HideUnchanged"] = "Hide unchanged",
                ["Collapse"] = "Hide inputs",
                ["Expand"] = "Show inputs",
                ["Comparison"] = "Comparison",
                ["KeyColumns"] = "Key columns",
                ["FormulaCachedValues"] = "Compare formula cached values",
                ["CompareDisplayText"] = "Compare display text",
                ["CompareStyles"] = "Compare cell styles",
                ["CompareRowMetadata"] = "Compare row metadata",
                ["BlankAsMissing"] = "Treat explicit blanks as missing",
                ["Save"] = "Save",
                ["SettingsSaved"] = "Settings saved",
                ["InvalidKeyColumns"] = "Enter key columns as letters or positive numbers.",
                ["ComparisonGrid"] = "Synchronized comparison grid",
                ["LegacyWorkbookUnsupported"] = "Legacy .xls workbooks are not supported. Convert the file to .xlsx first.",
            },
            ["zh-CN"] = new Dictionary<string, string>
            {
                ["AppTitle"] = "ExcelMerge",
                ["Compare"] = "比较",
                ["Merge"] = "合并",
                ["Open"] = "运行",
                ["Cancel"] = "取消",
                ["SaveResult"] = "保存结果",
                ["Settings"] = "设置",
                ["Diagnostics"] = "诊断",
                ["Base"] = "基准",
                ["Local"] = "本地",
                ["Remote"] = "远端",
                ["Browse"] = "浏览",
                ["Sheets"] = "工作表",
                ["Recent"] = "最近使用",
                ["Search"] = "搜索",
                ["Exact"] = "完全匹配",
                ["CaseSensitive"] = "区分大小写",
                ["Regex"] = "正则表达式",
                ["Previous"] = "上一个",
                ["Next"] = "下一个",
                ["Changes"] = "更改",
                ["Conflicts"] = "冲突",
                ["Inspector"] = "检查器",
                ["Result"] = "结果",
                ["Formula"] = "公式",
                ["Type"] = "类型",
                ["Value"] = "值",
                ["UseLocal"] = "使用本地",
                ["UseRemote"] = "使用远端",
                ["UseBoth"] = "保留两者",
                ["UseCustom"] = "自定义",
                ["CustomValue"] = "自定义值",
                ["CopyTsv"] = "复制 TSV",
                ["CopyCsv"] = "复制 CSV",
                ["Ready"] = "就绪",
                ["NoSession"] = "选择文件并运行比较。",
                ["Language"] = "语言",
                ["Theme"] = "主题",
                ["Light"] = "浅色",
                ["Dark"] = "深色",
                ["System"] = "跟随系统",
                ["Grid"] = "网格",
                ["RowHeight"] = "行高",
                ["ColumnWidth"] = "列宽",
                ["CacheRows"] = "缓存行数",
                ["Close"] = "关闭",
                ["OperationLog"] = "操作日志",
                ["Unresolved"] = "未解决",
                ["Resolved"] = "已解决",
                ["Loading"] = "正在加载工作簿",
                ["Saving"] = "正在保存结果",
                ["Completed"] = "已完成",
                ["Failed"] = "操作失败",
                ["Swap"] = "交换两侧",
                ["HideUnchanged"] = "隐藏未更改行",
                ["Collapse"] = "隐藏输入",
                ["Expand"] = "显示输入",
                ["Comparison"] = "比较",
                ["KeyColumns"] = "关键列",
                ["FormulaCachedValues"] = "比较公式缓存值",
                ["CompareDisplayText"] = "比较显示文本",
                ["CompareStyles"] = "比较单元格样式",
                ["CompareRowMetadata"] = "比较行元数据",
                ["BlankAsMissing"] = "将显式空白视为缺失",
                ["Save"] = "保存",
                ["SettingsSaved"] = "设置已保存",
                ["InvalidKeyColumns"] = "请使用列字母或正整数输入关键列。",
                ["ComparisonGrid"] = "同步比较网格",
                ["LegacyWorkbookUnsupported"] = "不支持旧版 .xls 工作簿。请先将文件转换为 .xlsx。",
            },
        };

    public static string CurrentCulture { get; private set; } = "en-US";

    public static string Get(string key) =>
        Cultures.TryGetValue(CurrentCulture, out var values) && values.TryGetValue(key, out var value)
            ? value
            : Cultures["en-US"].TryGetValue(key, out value) ? value : key;

    public static void Apply(string culture)
    {
        if (!Cultures.ContainsKey(culture))
        {
            culture = "en-US";
        }

        CurrentCulture = culture;
        if (Avalonia.Application.Current is not { } application)
        {
            return;
        }

        foreach (var pair in Cultures["en-US"])
        {
            application.Resources[pair.Key] = Get(pair.Key);
        }
    }
}
