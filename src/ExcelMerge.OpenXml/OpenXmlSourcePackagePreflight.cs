using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlSourcePackagePreflight
{
    private const uint ZipLocalFileHeaderSignature = 0x04034B50;
    private const uint ZipCentralDirectoryHeaderSignature = 0x02014B50;
    private const uint ZipEndOfCentralDirectorySignature = 0x06054B50;
    private const ulong CompoundFileSignature = 0xE11AB1A1E011CFD0;

    public static void Validate(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateContainerHeader(path);
        try
        {
            using var document = SpreadsheetDocument.Open(
                path,
                isEditable: false,
                new OpenSettings { AutoSave = false });
            if (document.DigitalSignatureOriginPart is not null)
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    "Digitally signed workbooks cannot participate in a merge.",
                    path);
            }

            ValidateRelationshipGraph(document, path, cancellationToken);
        }
        catch (OpenXmlWriterException)
        {
            throw;
        }
        catch (Exception exception) when (exception is
            OpenXmlPackageException or
            InvalidDataException or
            FormatException)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                "A workbook source is not a valid Open XML package.",
                path,
                innerException: exception);
        }
    }

    private static void ValidateContainerHeader(string path)
    {
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read | FileShare.Delete,
            bufferSize: 8,
            FileOptions.SequentialScan);
        if (stream.Length < sizeof(ulong))
        {
            throw InvalidPackage(path, "The workbook source is not a ZIP-based Open XML package.");
        }

        Span<byte> header = stackalloc byte[sizeof(ulong)];
        stream.ReadExactly(header);
        if (BinaryPrimitives.ReadUInt64LittleEndian(header) == CompoundFileSignature)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Encrypted or compound-file workbooks cannot participate in a merge.",
                path);
        }

        if (BinaryPrimitives.ReadUInt32LittleEndian(header) != ZipLocalFileHeaderSignature)
        {
            throw InvalidPackage(path, "The workbook source is not a ZIP-based Open XML package.");
        }

        var generalPurposeFlags = BinaryPrimitives.ReadUInt16LittleEndian(header[6..]);
        if ((generalPurposeFlags & 0x0001) != 0)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Encrypted ZIP workbook entries cannot participate in a merge.",
                path);
        }

        if (HasEncryptedEntry(stream, path))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Encrypted ZIP workbook entries cannot participate in a merge.",
                path);
        }
    }

    private static bool HasEncryptedEntry(FileStream stream, string path)
    {
        const int endRecordLength = 22;
        const int maximumCommentLength = ushort.MaxValue;
        var tailLength = checked((int)Math.Min(
            stream.Length,
            endRecordLength + maximumCommentLength));
        var tail = new byte[tailLength];
        stream.Position = stream.Length - tailLength;
        stream.ReadExactly(tail);

        var endRecordOffset = -1;
        for (var index = tail.Length - endRecordLength; index >= 0; index--)
        {
            if (BinaryPrimitives.ReadUInt32LittleEndian(tail.AsSpan(index, sizeof(uint))) !=
                ZipEndOfCentralDirectorySignature)
            {
                continue;
            }

            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(
                tail.AsSpan(index + 20, sizeof(ushort)));
            if (index + endRecordLength + commentLength == tail.Length)
            {
                endRecordOffset = index;
                break;
            }
        }

        if (endRecordOffset < 0)
        {
            throw InvalidPackage(path, "The ZIP central directory is missing or invalid.");
        }

        var endRecord = tail.AsSpan(endRecordOffset, endRecordLength);
        var diskNumber = BinaryPrimitives.ReadUInt16LittleEndian(endRecord[4..]);
        var centralDirectoryDisk = BinaryPrimitives.ReadUInt16LittleEndian(endRecord[6..]);
        var entriesOnDisk = BinaryPrimitives.ReadUInt16LittleEndian(endRecord[8..]);
        var totalEntries = BinaryPrimitives.ReadUInt16LittleEndian(endRecord[10..]);
        var centralDirectorySize = BinaryPrimitives.ReadUInt32LittleEndian(endRecord[12..]);
        var centralDirectoryOffset = BinaryPrimitives.ReadUInt32LittleEndian(endRecord[16..]);
        if (diskNumber != 0 || centralDirectoryDisk != 0 || entriesOnDisk != totalEntries)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Multi-disk ZIP workbooks cannot participate in a merge.",
                path);
        }

        if (totalEntries == ushort.MaxValue ||
            centralDirectorySize == uint.MaxValue ||
            centralDirectoryOffset == uint.MaxValue)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "ZIP64 workbooks cannot yet participate in a merge.",
                path);
        }

        if ((ulong)centralDirectoryOffset + centralDirectorySize > (ulong)stream.Length)
        {
            throw InvalidPackage(path, "The ZIP central directory is outside the package.");
        }

        stream.Position = centralDirectoryOffset;
        var header = new byte[46];
        for (var index = 0; index < totalEntries; index++)
        {
            if (stream.Length - stream.Position < header.Length)
            {
                throw InvalidPackage(path, "A ZIP central-directory entry is truncated.");
            }

            stream.ReadExactly(header);
            if (BinaryPrimitives.ReadUInt32LittleEndian(header) !=
                ZipCentralDirectoryHeaderSignature)
            {
                throw InvalidPackage(path, "A ZIP central-directory entry is invalid.");
            }

            var flags = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(8));
            if ((flags & 0x0001) != 0)
            {
                return true;
            }

            var fileNameLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(28));
            var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(30));
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(32));
            var remainingLength = checked((long)fileNameLength + extraLength + commentLength);
            if (remainingLength > stream.Length - stream.Position)
            {
                throw InvalidPackage(path, "A ZIP central-directory entry has invalid lengths.");
            }

            stream.Position += remainingLength;
        }

        return false;
    }

    private static void ValidateRelationshipGraph(
        SpreadsheetDocument document,
        string path,
        CancellationToken cancellationToken)
    {
        var visited = new HashSet<OpenXmlPart>();
        var active = new HashSet<OpenXmlPart>();
        foreach (var pair in document.Parts)
        {
            Visit(pair.OpenXmlPart);
        }

        return;

        void Visit(OpenXmlPart part)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!active.Add(part))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPackage,
                    "A workbook source contains a cyclic package relationship graph.",
                    path);
            }

            if (!visited.Add(part))
            {
                active.Remove(part);
                return;
            }

            foreach (var child in part.Parts)
            {
                Visit(child.OpenXmlPart);
            }

            active.Remove(part);
        }
    }

    private static OpenXmlWriterException InvalidPackage(string path, string message) =>
        new(OpenXmlWriterError.InvalidPackage, message, path);
}
