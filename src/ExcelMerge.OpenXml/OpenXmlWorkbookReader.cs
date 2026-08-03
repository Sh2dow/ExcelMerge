using System.Buffers.Binary;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO.Compression;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.OpenXml;

/// <summary>A stateless, thread-safe streaming reader for standard, unencrypted <c>.xlsx</c> workbooks.</summary>
public sealed class OpenXmlWorkbookReader : IOpenXmlWorkbookReader
{
    private const uint ZipLocalFileHeaderSignature = 0x04034B50;
    private const ulong CompoundFileSignature = 0xE11AB1A1E011CFD0;

    public ValueTask<WorkbookMetadata> ReadMetadataAsync(
        string filePath,
        OpenXmlReaderOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var settings = ReaderSettings.Create(options);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var source = OpenAndPreflight(filePath, settings, cancellationToken);
            using var document = OpenDocument(source, settings);
            var discovered = DiscoverWorkbook(document, source, cancellationToken);
            source.ThrowIfChanged();
            return ValueTask.FromResult(discovered.Metadata);
        }
        catch (OpenXmlReaderException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (IsMalformedPackageException(exception))
        {
            throw InvalidPackage(filePath, exception);
        }
    }

    public async ValueTask<OpenXmlReaderResult> IndexAsync(
        string filePath,
        Workspace workspace,
        OpenXmlReaderOptions? options = null,
        IProgress<OpenXmlReaderProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentNullException.ThrowIfNull(workspace);
        var settings = ReaderSettings.Create(options);

        try
        {
            return await IndexCoreAsync(
                filePath,
                workspace,
                settings,
                progress,
                cancellationToken).ConfigureAwait(false);
        }
        catch (OpenXmlReaderException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (IsMalformedPackageException(exception))
        {
            throw InvalidPackage(filePath, exception);
        }
    }

    private static async ValueTask<OpenXmlReaderResult> IndexCoreAsync(
        string filePath,
        Workspace workspace,
        ReaderSettings settings,
        IProgress<OpenXmlReaderProgress>? progress,
        CancellationToken cancellationToken)
    {
        var createdStores = new List<ChunkedCellStore>();
        try
        {
            progress?.Report(new OpenXmlReaderProgress(OpenXmlReaderStage.Preflight, 0, 0));
            cancellationToken.ThrowIfCancellationRequested();

            using var source = OpenAndPreflight(filePath, settings, cancellationToken);
            using var document = OpenDocument(source, settings);
            var discovered = DiscoverWorkbook(document, source, cancellationToken);
            var selectedSheets = SelectSheets(discovered, settings, source.FullPath);

            progress?.Report(new OpenXmlReaderProgress(
                OpenXmlReaderStage.DiscoveringWorkbook,
                0,
                selectedSheets.Count));

            if (selectedSheets.Count == 0)
            {
                source.ThrowIfChanged();
                progress?.Report(new OpenXmlReaderProgress(OpenXmlReaderStage.Completed, 0, 0));
                return new OpenXmlReaderResult(
                    discovered.Metadata,
                    Array.Empty<OpenXmlReaderWorksheet>());
            }

            var styles = OpenXmlReaderStyleCatalog.Load(
                discovered.StylesPart,
                source.FullPath,
                settings.DisplayCulture,
                cancellationToken);
            using var sharedStrings = OpenXmlReaderSharedStrings.Load(
                discovered.SharedStringsPart,
                workspace,
                source.FullPath,
                count => progress?.Report(new OpenXmlReaderProgress(
                    OpenXmlReaderStage.ReadingSharedStrings,
                    0,
                    selectedSheets.Count,
                    SharedStringsRead: count)),
                cancellationToken);

            var indexedWorksheets = new OpenXmlReaderWorksheet[selectedSheets.Count];
            var updatedMetadata = discovered.Metadata.Sheets.ToArray();

            for (var index = 0; index < selectedSheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var sheet = selectedSheets[index];
                var store = workspace.CreateCellStore();
                createdStores.Add(store);

                progress?.Report(new OpenXmlReaderProgress(
                    OpenXmlReaderStage.IndexingWorksheet,
                    index,
                    selectedSheets.Count,
                    sheet.Metadata.Id,
                    sheet.Metadata.Name));

                var indexed = await OpenXmlReaderWorksheetIndexer.IndexAsync(
                    sheet,
                    store,
                    sharedStrings,
                    styles,
                    discovered.Metadata.Uses1904DateSystem,
                    settings,
                    (rows, cells) => progress?.Report(new OpenXmlReaderProgress(
                        OpenXmlReaderStage.IndexingWorksheet,
                        index,
                        selectedSheets.Count,
                        sheet.Metadata.Id,
                        sheet.Metadata.Name,
                        rows,
                        cells,
                        sharedStrings.Count)),
                    cancellationToken).ConfigureAwait(false);

                await store.FlushAsync(cancellationToken).ConfigureAwait(false);
                updatedMetadata[sheet.Metadata.Position] = indexed.Metadata;
                indexedWorksheets[index] = new OpenXmlReaderWorksheet(
                    indexed.Metadata,
                    store,
                    indexed.StoredRowIndices);

                progress?.Report(new OpenXmlReaderProgress(
                    OpenXmlReaderStage.IndexingWorksheet,
                    index + 1,
                    selectedSheets.Count,
                    sheet.Metadata.Id,
                    sheet.Metadata.Name,
                    indexed.StoredRowIndices.LongLength,
                    indexed.CellCount,
                    sharedStrings.Count));
            }

            source.ThrowIfChanged();
            var resultMetadata = discovered.Metadata with
            {
                Sheets = new ReadOnlyCollection<SheetMetadata>(updatedMetadata),
            };

            progress?.Report(new OpenXmlReaderProgress(
                OpenXmlReaderStage.Completed,
                selectedSheets.Count,
                selectedSheets.Count,
                SharedStringsRead: sharedStrings.Count));

            return new OpenXmlReaderResult(
                resultMetadata,
                OpenXmlReaderWorksheet.AsReadOnly(indexedWorksheets));
        }
        catch
        {
            await DisposeFailedStoresAsync(createdStores).ConfigureAwait(false);
            throw;
        }
    }

    private static SourceFile OpenAndPreflight(
        string filePath,
        ReaderSettings settings,
        CancellationToken cancellationToken)
    {
        var fullPath = Path.GetFullPath(filePath);
        var extension = Path.GetExtension(fullPath);
        if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.LegacyBinaryWorkbook,
                "Legacy binary .xls workbooks are not supported.",
                fullPath);
        }

        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.UnsupportedFileExtension,
                "Only standard .xlsx workbook files are supported.",
                fullPath);
        }

        var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 64 * 1024,
            FileOptions.SequentialScan);

        try
        {
            var source = new SourceFile(fullPath, stream);
            PreflightZip(source, settings, cancellationToken);
            return source;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private static void PreflightZip(
        SourceFile source,
        ReaderSettings settings,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Length < sizeof(ulong))
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.NotZipPackage,
                "The input is not a ZIP-based Open XML package.",
                source.FullPath);
        }

        Span<byte> header = stackalloc byte[sizeof(ulong)];
        source.Stream.Position = 0;
        source.Stream.ReadExactly(header);
        var signature = BinaryPrimitives.ReadUInt64LittleEndian(header);
        if (signature == CompoundFileSignature)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.EncryptedPackage,
                "The input is an encrypted or compound-file workbook and cannot be read as .xlsx.",
                source.FullPath);
        }

        if (BinaryPrimitives.ReadUInt32LittleEndian(header) != ZipLocalFileHeaderSignature)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.NotZipPackage,
                "The input is not a ZIP-based Open XML package.",
                source.FullPath);
        }

        var generalPurposeFlags = BinaryPrimitives.ReadUInt16LittleEndian(header[6..]);
        if ((generalPurposeFlags & 0x0001) != 0)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.EncryptedPackage,
                "Encrypted ZIP entries are not supported.",
                source.FullPath);
        }

        source.Stream.Position = 0;
        try
        {
            using var archive = new ZipArchive(
                source.Stream,
                ZipArchiveMode.Read,
                leaveOpen: true,
                entryNameEncoding: null);
            var entryNames = new HashSet<string>(StringComparer.Ordinal);
            var hasContentTypes = false;
            var hasRootRelationships = false;
            long uncompressedBytes = 0;

            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateEntryName(entry.FullName, source.FullPath);
                if (!entryNames.Add(entry.FullName))
                {
                    throw new OpenXmlReaderException(
                        OpenXmlReaderError.InvalidPackage,
                        $"The package contains a duplicate ZIP entry named '{entry.FullName}'.",
                        source.FullPath);
                }

                hasContentTypes |= entry.FullName == "[Content_Types].xml";
                hasRootRelationships |= entry.FullName == "_rels/.rels";

                try
                {
                    uncompressedBytes = checked(uncompressedBytes + entry.Length);
                }
                catch (OverflowException exception)
                {
                    throw new OpenXmlReaderException(
                        OpenXmlReaderError.InvalidPackage,
                        "The package's aggregate uncompressed size is invalid.",
                        source.FullPath,
                        innerException: exception);
                }

                if (settings.MaximumUncompressedPackageBytes > 0 &&
                    uncompressedBytes > settings.MaximumUncompressedPackageBytes)
                {
                    throw new OpenXmlReaderException(
                        OpenXmlReaderError.InvalidPackage,
                        "The package exceeds the configured uncompressed-size limit.",
                        source.FullPath);
                }
            }

            if (!hasContentTypes || !hasRootRelationships)
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.InvalidPackage,
                    "The ZIP file is missing required Open Packaging Convention metadata.",
                    source.FullPath);
            }
        }
        catch (OpenXmlReaderException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidPackage,
                "The ZIP package is corrupt or uses unsupported compression.",
                source.FullPath,
                innerException: exception);
        }
        finally
        {
            source.Stream.Position = 0;
        }
    }

    private static SpreadsheetDocument OpenDocument(SourceFile source, ReaderSettings settings)
    {
        SpreadsheetDocument document;
        try
        {
            document = SpreadsheetDocument.Open(
                source.Stream,
                isEditable: false,
                new OpenSettings
                {
                    AutoSave = false,
                    MaxCharactersInPart = settings.MaximumCharactersInPart,
                });
        }
        catch (Exception exception) when (IsMalformedPackageException(exception))
        {
            throw InvalidPackage(source.FullPath, exception);
        }

        var documentType = document.DocumentType;
        if (documentType != SpreadsheetDocumentType.Workbook)
        {
            document.Dispose();
            throw new OpenXmlReaderException(
                OpenXmlReaderError.UnsupportedWorkbookType,
                $"Spreadsheet document type '{documentType}' is not a standard .xlsx workbook.",
                source.FullPath);
        }

        return document;
    }

    private static DiscoveredWorkbook DiscoverWorkbook(
        SpreadsheetDocument document,
        SourceFile source,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.MissingWorkbookPart,
                "The package does not contain a workbook part.",
                source.FullPath);
        }

        var workbook = workbookPart.Workbook;
        if (workbook is null)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidPackage,
                "The workbook part does not contain a workbook root element.",
                source.FullPath);
        }

        var sheetsElement = workbook.Sheets;
        if (sheetsElement is null)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidPackage,
                "The workbook does not declare any worksheets.",
                source.FullPath);
        }

        var discoveredSheets = new List<DiscoveredSheet>();
        var relationshipIds = new HashSet<string>(StringComparer.Ordinal);
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetIds = new HashSet<uint>();
        var partUris = new HashSet<string>(StringComparer.Ordinal);
        var position = 0;

        foreach (var sheet in sheetsElement.Elements<Sheet>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relationshipId = sheet.Id?.Value;
            var name = sheet.Name?.Value;
            var numericSheetId = sheet.SheetId?.Value;

            if (string.IsNullOrWhiteSpace(relationshipId) ||
                string.IsNullOrWhiteSpace(name) ||
                numericSheetId is null or 0 ||
                !relationshipIds.Add(relationshipId) ||
                !sheetNames.Add(name) ||
                !sheetIds.Add(numericSheetId.Value))
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.InvalidRelationship,
                    "The workbook contains missing or duplicate sheet identifiers or names.",
                    source.FullPath,
                    relationshipId);
            }

            if (!workbookPart.TryGetPartById(relationshipId, out var relatedPart))
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.InvalidRelationship,
                    $"Worksheet '{name}' refers to missing or external relationship '{relationshipId}'.",
                    source.FullPath,
                    relationshipId);
            }

            if (relatedPart is not WorksheetPart worksheetPart)
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.UnsupportedSheetType,
                    $"Sheet '{name}' does not refer to a worksheet part.",
                    source.FullPath,
                    relationshipId);
            }

            if (!partUris.Add(worksheetPart.Uri.OriginalString))
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.InvalidRelationship,
                    "Multiple sheet declarations refer to the same worksheet part.",
                    source.FullPath,
                    relationshipId);
            }

            var sheetMetadata = new SheetMetadata(
                relationshipId,
                name,
                position,
                GetVisibility(sheet, source.FullPath, relationshipId));
            discoveredSheets.Add(new DiscoveredSheet(sheetMetadata, worksheetPart, source.FullPath));
            position++;
        }

        if (discoveredSheets.Count == 0)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidPackage,
                "The workbook does not contain a worksheet.",
                source.FullPath);
        }

        var stylesPart = GetOptionalSingletonPart<WorkbookStylesPart>(
            workbookPart,
            "styles",
            source.FullPath);
        var sharedStringsPart = GetOptionalSingletonPart<SharedStringTablePart>(
            workbookPart,
            "shared strings",
            source.FullPath);
        var metadataSheets = new ReadOnlyCollection<SheetMetadata>(
            discoveredSheets.Select(static sheet => sheet.Metadata).ToArray());
        var metadata = new WorkbookMetadata(
            source.DisplayName,
            WorkbookFormat.Xlsx,
            metadataSheets,
            source.FullPath,
            source.Length,
            source.LastModifiedUtc,
            workbook.WorkbookProperties?.Date1904?.Value ?? false);

        return new DiscoveredWorkbook(
            metadata,
            discoveredSheets,
            stylesPart,
            sharedStringsPart);
    }

    private static SheetVisibility GetVisibility(Sheet sheet, string sourcePath, string relationshipId)
    {
        var state = sheet.State?.Value;
        if (state is null || state == SheetStateValues.Visible)
        {
            return SheetVisibility.Visible;
        }

        if (state == SheetStateValues.Hidden)
        {
            return SheetVisibility.Hidden;
        }

        if (state == SheetStateValues.VeryHidden)
        {
            return SheetVisibility.VeryHidden;
        }

        throw new OpenXmlReaderException(
            OpenXmlReaderError.InvalidPackage,
            "A worksheet has an invalid visibility value.",
            sourcePath,
            relationshipId);
    }

    private static TPart? GetOptionalSingletonPart<TPart>(
        WorkbookPart workbookPart,
        string description,
        string sourcePath)
        where TPart : OpenXmlPart
    {
        var parts = workbookPart.GetPartsOfType<TPart>().Take(2).ToArray();
        if (parts.Length > 1)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidRelationship,
                $"The workbook has multiple {description} relationships.",
                sourcePath);
        }

        return parts.Length == 0 ? null : parts[0];
    }

    private static List<DiscoveredSheet> SelectSheets(
        DiscoveredWorkbook workbook,
        ReaderSettings settings,
        string sourcePath)
    {
        if (!settings.HasSheetSelection)
        {
            return [.. workbook.Sheets];
        }

        foreach (var requestedId in settings.SheetIds)
        {
            if (!workbook.Sheets.Any(sheet => sheet.Metadata.Id == requestedId))
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.SheetNotFound,
                    $"The workbook does not contain sheet relationship '{requestedId}'.",
                    sourcePath,
                    requestedId);
            }
        }

        foreach (var requestedName in settings.SheetNames)
        {
            if (!workbook.Sheets.Any(sheet =>
                string.Equals(sheet.Metadata.Name, requestedName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.SheetNotFound,
                    $"The workbook does not contain a worksheet named '{requestedName}'.",
                    sourcePath);
            }
        }

        return workbook.Sheets
            .Where(sheet =>
                settings.SheetIds.Contains(sheet.Metadata.Id) ||
                settings.SheetNames.Contains(sheet.Metadata.Name))
            .ToList();
    }

    private static void ValidateEntryName(string entryName, string sourcePath)
    {
        if (string.IsNullOrEmpty(entryName) ||
            entryName.StartsWith('/') ||
            entryName.Contains('\\') ||
            entryName.Contains('\0'))
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidPackage,
                "The package contains an invalid ZIP entry name.",
                sourcePath);
        }

        foreach (var segment in entryName.Split('/'))
        {
            if (segment is "." or "..")
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.InvalidPackage,
                    "The package contains a ZIP entry that escapes its package path.",
                    sourcePath);
            }
        }
    }

    private static async ValueTask DisposeFailedStoresAsync(List<ChunkedCellStore> stores)
    {
        for (var index = stores.Count - 1; index >= 0; index--)
        {
            var store = stores[index];
            var directoryPath = store.DirectoryPath;
            try
            {
                await store.DisposeAsync().ConfigureAwait(false);
                if (Directory.Exists(directoryPath))
                {
                    Directory.Delete(directoryPath, recursive: true);
                }
            }
            catch
            {
                // Preserve the indexing exception; workspace disposal owns final cleanup.
            }
        }
    }

    private static bool IsMalformedPackageException(Exception exception) => exception is
        OpenXmlPackageException or
        FileFormatException or
        InvalidDataException or
        XmlException or
        FormatException;

    private static OpenXmlReaderException InvalidPackage(string filePath, Exception exception) =>
        new(
            OpenXmlReaderError.InvalidPackage,
            "The input is not a valid .xlsx Open XML package.",
            Path.GetFullPath(filePath),
            innerException: exception);

    internal sealed record DiscoveredSheet(
        SheetMetadata Metadata,
        WorksheetPart WorksheetPart,
        string SourcePath);

    private sealed record DiscoveredWorkbook(
        WorkbookMetadata Metadata,
        IReadOnlyList<DiscoveredSheet> Sheets,
        WorkbookStylesPart? StylesPart,
        SharedStringTablePart? SharedStringsPart);

    internal sealed class ReaderSettings
    {
        private ReaderSettings(
            HashSet<string> sheetIds,
            HashSet<string> sheetNames,
            bool hasSheetSelection,
            bool includeDisplayText,
            CultureInfo displayCulture,
            int progressIntervalRows,
            long maximumUncompressedPackageBytes,
            long maximumCharactersInPart)
        {
            SheetIds = sheetIds;
            SheetNames = sheetNames;
            HasSheetSelection = hasSheetSelection;
            IncludeDisplayText = includeDisplayText;
            DisplayCulture = displayCulture;
            ProgressIntervalRows = progressIntervalRows;
            MaximumUncompressedPackageBytes = maximumUncompressedPackageBytes;
            MaximumCharactersInPart = maximumCharactersInPart;
        }

        public HashSet<string> SheetIds { get; }

        public HashSet<string> SheetNames { get; }

        public bool HasSheetSelection { get; }

        public bool IncludeDisplayText { get; }

        public CultureInfo DisplayCulture { get; }

        public int ProgressIntervalRows { get; }

        public long MaximumUncompressedPackageBytes { get; }

        public long MaximumCharactersInPart { get; }

        public static ReaderSettings Create(OpenXmlReaderOptions? options)
        {
            options ??= new OpenXmlReaderOptions();
            ArgumentNullException.ThrowIfNull(options.DisplayCulture);
            if (options.ProgressIntervalRows <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    "The progress interval must be positive.");
            }

            if (options.MaximumUncompressedPackageBytes < 0 || options.MaximumCharactersInPart < 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    "Package limits cannot be negative.");
            }

            var ids = CreateSelectionSet(options.SheetIds, StringComparer.Ordinal, "sheet identifier");
            var names = CreateSelectionSet(
                options.SheetNames,
                StringComparer.OrdinalIgnoreCase,
                "sheet name");
            var culture = CultureInfo.ReadOnly((CultureInfo)options.DisplayCulture.Clone());

            return new ReaderSettings(
                ids,
                names,
                options.SheetIds is not null || options.SheetNames is not null,
                options.IncludeDisplayText,
                culture,
                options.ProgressIntervalRows,
                options.MaximumUncompressedPackageBytes,
                options.MaximumCharactersInPart);
        }

        private static HashSet<string> CreateSelectionSet(
            IReadOnlyCollection<string>? values,
            StringComparer comparer,
            string description)
        {
            var result = new HashSet<string>(comparer);
            if (values is null)
            {
                return result;
            }

            foreach (var value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    throw new ArgumentException($"A selected {description} cannot be empty.", nameof(values));
                }

                result.Add(value);
            }

            return result;
        }
    }

    private sealed class SourceFile : IDisposable
    {
        public SourceFile(string fullPath, FileStream stream)
        {
            FullPath = fullPath;
            Stream = stream;
            var fileInfo = new FileInfo(fullPath);
            DisplayName = fileInfo.Name;
            Length = stream.Length;
            LastWriteTimeUtc = fileInfo.LastWriteTimeUtc;
            LastModifiedUtc = new DateTimeOffset(LastWriteTimeUtc);
        }

        public string FullPath { get; }

        public string DisplayName { get; }

        public FileStream Stream { get; }

        public long Length { get; }

        public DateTime LastWriteTimeUtc { get; }

        public DateTimeOffset LastModifiedUtc { get; }

        public void ThrowIfChanged()
        {
            var fileInfo = new FileInfo(FullPath);
            fileInfo.Refresh();
            if (!fileInfo.Exists ||
                fileInfo.Length != Length ||
                fileInfo.LastWriteTimeUtc != LastWriteTimeUtc)
            {
                throw new OpenXmlReaderException(
                    OpenXmlReaderError.SourceChanged,
                    "The source workbook changed while it was being read.",
                    FullPath);
            }
        }

        public void Dispose() => Stream.Dispose();
    }
}
