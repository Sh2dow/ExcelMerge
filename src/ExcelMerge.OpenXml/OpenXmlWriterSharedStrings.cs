using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Storage;
using System.Xml;

namespace ExcelMerge.OpenXml;

internal sealed class OpenXmlWriterSharedStrings : IDisposable
{
    private readonly ChunkedTextStore? _store;

    private OpenXmlWriterSharedStrings(ChunkedTextStore? store)
    {
        _store = store;
    }

    public long Count => _store?.Count ?? 0;

    public static OpenXmlWriterSharedStrings Load(
        SharedStringTablePart? part,
        string storeDirectory,
        CancellationToken cancellationToken)
    {
        if (part is null)
        {
            return new OpenXmlWriterSharedStrings(null);
        }

        var store = new ChunkedTextStore(storeDirectory);
        try
        {
            var sawRoot = false;
            using var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(part);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.IsStartElement)
                {
                    continue;
                }

                if (reader.Depth == 0)
                {
                    sawRoot = reader.ElementType == typeof(SharedStringTable);
                    if (!sawRoot)
                    {
                        throw InvalidTable("The REMOTE shared-string part has an invalid root element.");
                    }
                }

                if (reader.ElementType != typeof(SharedStringItem))
                {
                    continue;
                }

                if (reader.LoadCurrentElement() is not SharedStringItem item)
                {
                    throw InvalidTable("A REMOTE shared-string item could not be read.");
                }

                store.Append(item.OuterXml, cancellationToken);
            }

            if (!sawRoot)
            {
                throw InvalidTable("The REMOTE shared-string part is empty.");
            }

            store.Flush();
            return new OpenXmlWriterSharedStrings(store);
        }
        catch
        {
            store.Dispose();
            throw;
        }
    }

    public SharedStringItem Read(int index, string? cellReference)
    {
        if (index < 0 || index >= Count || _store is null)
        {
            throw InvalidTable(
                $"REMOTE cell '{cellReference}' has an invalid shared-string reference.");
        }

        try
        {
            return new SharedStringItem(_store.Read(index));
        }
        catch (OpenXmlWriterException)
        {
            throw;
        }
        catch (Exception exception) when (exception is
            InvalidDataException or
            FormatException or
            XmlException or
            OpenXmlPackageException)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"REMOTE shared-string item {index} is invalid.",
                innerException: exception);
        }
    }

    public void Dispose() => _store?.Dispose();

    private static OpenXmlWriterException InvalidTable(string message) =>
        new(OpenXmlWriterError.InvalidPackage, message);
}
