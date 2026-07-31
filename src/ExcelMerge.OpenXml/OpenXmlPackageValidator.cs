using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlPackageValidator
{
    private const long LargePartThresholdBytes = 16L * 1024 * 1024;

    public static void Validate(
        string packagePath,
        int maximumErrors,
        CancellationToken cancellationToken)
    {
        var largeParts = FindLargeParts(packagePath, cancellationToken);
        if (!largeParts.HasAny)
        {
            ValidateCompletePackage(packagePath, packagePath, maximumErrors, cancellationToken);
            return;
        }

        var streamedErrors = ValidateLargeParts(
            packagePath,
            largeParts,
            maximumErrors,
            cancellationToken);
        ThrowIfErrors(packagePath, streamedErrors);

        var shadowPath = CreateTemporaryPath(packagePath);
        try
        {
            File.Copy(packagePath, shadowPath, overwrite: false);
            StripLargeCollections(shadowPath, largeParts, cancellationToken);
            ValidateCompletePackage(shadowPath, packagePath, maximumErrors, cancellationToken);
        }
        finally
        {
            try
            {
                if (File.Exists(shadowPath))
                    File.Delete(shadowPath);
            }
            catch
            {
                // Preserve the validation result; startup cleanup handles an abandoned shadow copy.
            }
        }
    }

    private static LargePartSet FindLargeParts(
        string packagePath,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(
            packagePath,
            isEditable: false,
            new OpenSettings { AutoSave = false });
        var workbookPart = document.WorkbookPart ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.ValidationFailed,
            "The generated workbook has no workbook part.",
            packagePath);
        var worksheets = new HashSet<string>(StringComparer.Ordinal);
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.Length > LargePartThresholdBytes)
                worksheets.Add(worksheetPart.Uri.OriginalString);
        }

        string? sharedStrings = null;
        if (workbookPart.SharedStringTablePart is { } sharedStringPart)
        {
            using var stream = sharedStringPart.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.Length > LargePartThresholdBytes)
                sharedStrings = sharedStringPart.Uri.OriginalString;
        }

        return new LargePartSet(worksheets, sharedStrings);
    }

    private static OpenXmlValidationError[] ValidateLargeParts(
        string packagePath,
        LargePartSet largeParts,
        int maximumErrors,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(
            packagePath,
            isEditable: false,
            new OpenSettings { AutoSave = false });
        var workbookPart = document.WorkbookPart ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.ValidationFailed,
            "The generated workbook has no workbook part.",
            packagePath);
        var validator = CreateValidator(maximumErrors);
        var errors = new List<OpenXmlValidationError>();
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!largeParts.Worksheets.Contains(worksheetPart.Uri.OriginalString))
                continue;

            ValidateCollectionItems<Row>(
                worksheetPart,
                typeof(Worksheet),
                validator,
                errors,
                maximumErrors,
                cancellationToken);
            if (errors.Count >= maximumErrors)
                return [.. errors];
        }

        if (largeParts.SharedStrings is not null &&
            workbookPart.SharedStringTablePart is { } sharedStringPart)
        {
            ValidateCollectionItems<SharedStringItem>(
                sharedStringPart,
                typeof(SharedStringTable),
                validator,
                errors,
                maximumErrors,
                cancellationToken);
        }

        return [.. errors];
    }

    private static void ValidateCollectionItems<TElement>(
        OpenXmlPart part,
        Type expectedRootType,
        OpenXmlValidator validator,
        List<OpenXmlValidationError> errors,
        int maximumErrors,
        CancellationToken cancellationToken)
        where TElement : OpenXmlElement
    {
        var sawRoot = false;
        using var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(part);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.IsStartElement)
                continue;

            if (reader.Depth == 0)
            {
                sawRoot = reader.ElementType == expectedRootType;
                if (!sawRoot)
                {
                    errors.Add(new OpenXmlValidationError(
                        "Sch_InvalidRoot",
                        $"Part '{part.Uri}' has an invalid root element.",
                        part.Uri.OriginalString,
                        null));
                    return;
                }
            }

            if (reader.ElementType != typeof(TElement))
                continue;

            if (reader.LoadCurrentElement() is not TElement element)
            {
                errors.Add(new OpenXmlValidationError(
                    "Sch_InvalidElement",
                    $"An element in part '{part.Uri}' could not be read.",
                    part.Uri.OriginalString,
                    null));
                return;
            }

            foreach (var error in validator.Validate(element, cancellationToken))
            {
                errors.Add(MapError(error, part.Uri.OriginalString));
                if (errors.Count >= maximumErrors)
                    return;
            }
        }

        if (!sawRoot)
        {
            errors.Add(new OpenXmlValidationError(
                "Sch_MissingRoot",
                $"Part '{part.Uri}' is empty.",
                part.Uri.OriginalString,
                null));
        }
    }

    private static void StripLargeCollections(
        string shadowPath,
        LargePartSet largeParts,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(
            shadowPath,
            isEditable: true,
            new OpenSettings { AutoSave = false });
        var workbookPart = document.WorkbookPart ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.ValidationFailed,
            "The generated workbook has no workbook part.",
            shadowPath);
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (largeParts.Worksheets.Contains(worksheetPart.Uri.OriginalString))
                StripElements<Row>(worksheetPart, shadowPath, cancellationToken);
        }

        if (largeParts.SharedStrings is not null &&
            workbookPart.SharedStringTablePart is { } sharedStringPart)
        {
            StripElements<SharedStringItem>(sharedStringPart, shadowPath, cancellationToken);
        }
    }

    private static void StripElements<TElement>(
        OpenXmlPart part,
        string shadowPath,
        CancellationToken cancellationToken)
        where TElement : OpenXmlElement
    {
        var directory = Path.GetDirectoryName(shadowPath)!;
        var scratchPath = Path.Combine(directory, $".{Path.GetFileName(shadowPath)}.{Guid.NewGuid():N}.xml");
        try
        {
            using (var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(part, readMiscNodes: true))
            using (var stream = new FileStream(
                scratchPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 128 * 1024,
                FileOptions.SequentialScan))
            using (var writer = DocumentFormat.OpenXml.OpenXmlWriter.Create(stream))
            {
                writer.WriteStartDocument();
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.IsStartElement && reader.ElementType == typeof(TElement))
                    {
                        if (reader.LoadCurrentElement() is null)
                            throw new InvalidDataException($"Part '{part.Uri}' contains an unreadable element.");
                    }
                    else if (reader.IsStartElement)
                    {
                        writer.WriteStartElement(reader);
                    }
                    else if (reader.IsEndElement)
                    {
                        writer.WriteEndElement();
                    }
                    else if (reader.IsMiscNode)
                    {
                        writer.WriteString(reader.GetText());
                    }
                }
            }

            using var input = new FileStream(
                scratchPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 128 * 1024,
                FileOptions.SequentialScan);
            part.FeedData(input);
            part.UnloadRootElement();
        }
        finally
        {
            try
            {
                if (File.Exists(scratchPath))
                    File.Delete(scratchPath);
            }
            catch
            {
                // The shadow package is never committed, so cleanup can retry with its parent.
            }
        }
    }

    private static void ValidateCompletePackage(
        string validationPath,
        string reportedPath,
        int maximumErrors,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(
            validationPath,
            isEditable: false,
            new OpenSettings { AutoSave = false });
        var errors = CreateValidator(maximumErrors)
            .Validate(document, cancellationToken)
            .Take(maximumErrors)
            .Select(error => MapError(error, error.Part?.Uri.OriginalString))
            .ToArray();
        ThrowIfErrors(reportedPath, errors);
    }

    private static OpenXmlValidator CreateValidator(int maximumErrors) =>
        new(FileFormatVersions.Microsoft365)
        {
            MaxNumberOfErrors = maximumErrors,
        };

    private static OpenXmlValidationError MapError(
        ValidationErrorInfo error,
        string? fallbackPartUri) =>
        new(
            error.Id,
            error.Description ?? "Open XML validation failed.",
            error.Part?.Uri.OriginalString ?? fallbackPartUri,
            error.Path?.XPath);

    private static void ThrowIfErrors(
        string packagePath,
        IReadOnlyCollection<OpenXmlValidationError> errors)
    {
        if (errors.Count == 0)
            return;

        throw new OpenXmlWriterException(
            OpenXmlWriterError.ValidationFailed,
            $"The generated workbook contains {errors.Count} Open XML validation error(s).",
            packagePath,
            errors.ToArray());
    }

    private static string CreateTemporaryPath(string packagePath)
    {
        var directory = Path.GetDirectoryName(packagePath)!;
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var candidate = Path.Combine(
                directory,
                $".{Path.GetFileName(packagePath)}.{Guid.NewGuid():N}.validation.tmp");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
                return candidate;
        }

        throw new IOException("A validation shadow path could not be allocated.");
    }

    private sealed record LargePartSet(
        IReadOnlySet<string> Worksheets,
        string? SharedStrings)
    {
        public bool HasAny => Worksheets.Count != 0 || SharedStrings is not null;
    }
}
