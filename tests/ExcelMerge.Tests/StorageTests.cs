using ExcelMerge.Domain;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class StorageTests
{
    [TestMethod]
    public async Task Workspace_owns_stores_and_removes_its_directory()
    {
        using var temporaryDirectory = new TestDirectory();
        string workspacePath;
        ChunkedCellStore store;

        await using (var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
            DirectoryPrefix = "workspace-",
        }))
        {
            workspacePath = workspace.DirectoryPath;
            store = workspace.CreateCellStore("sheet");

            Assert.IsTrue(Directory.Exists(workspacePath));
            Assert.IsTrue(File.Exists(store.IndexFilePath));
            Assert.AreEqual(0L, store.RowCount);
            Assert.AreEqual(1, store.ChunkCount);
            Assert.AreEqual(16L, store.DataLengthBytes);
            Assert.AreEqual(16L, store.IndexLengthBytes);
        }

        Assert.IsTrue(store.IsDisposed);
        Assert.IsFalse(Directory.Exists(workspacePath));
    }

    [TestMethod]
    public async Task Cell_store_round_trips_all_typed_values_and_row_metadata()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
            RowCacheCapacity = 4,
            RowCacheByteLimit = 1024 * 1024,
        });
        var store = workspace.CreateCellStore();
        var timestamp = new DateTime(2026, 7, 31, 12, 34, 56, DateTimeKind.Utc);
        var row = new RowRecord(
            7,
            new CellRecord[]
            {
                new(new CellAddress(7, 0), CellValue.Blank),
                new(new CellAddress(7, 1), CellValue.FromText("text", "shown"), 1),
                new(new CellAddress(7, 2), CellValue.FromNumber(-12.5)),
                new(new CellAddress(7, 3), CellValue.FromBoolean(true)),
                new(new CellAddress(7, 4), CellValue.FromDateTime(timestamp)),
                new(new CellAddress(7, 5), CellValue.FromError("#N/A")),
                new(new CellAddress(7, 6), CellValue.FromFormula("A1+B1")),
                new(
                    new CellAddress(7, 7),
                    CellValue.FromFormula("SUM(A1:B1)", CellScalar.FromNumber(3), "3"),
                    4),
                new(
                    new CellAddress(7, 8),
                    CellValue.FromFormula("\"\"", CellScalar.Blank)),
            },
            height: 22.5,
            isHidden: true,
            styleIndex: 3);

        var offset = await store.AppendRowAsync(row);
        var actual = await store.ReadRowAsync(offset);

        Assert.AreEqual(0L, offset.RowOrdinal);
        Assert.AreEqual(16L, offset.ByteOffset);
        Assert.AreEqual(row.CellCount, offset.CellCount);
        Assert.AreEqual(row.RowIndex, offset.SourceRowIndex);
        Assert.AreEqual(1L, store.RowCount);
        Assert.AreEqual(40L, store.IndexLengthBytes);
        TestData.AssertRowsEqual(row, actual);
        Assert.IsTrue(store.TryGetCachedRow(0, out var cached));
        TestData.AssertRowsEqual(row, cached);

        var mismatchedOffset = new RowOffset(
            offset.RowOrdinal,
            offset.ChunkNumber,
            offset.ByteOffset,
            offset.ByteLength,
            offset.CellCount,
            offset.SourceRowIndex + 1);
        await Assert.ThrowsExactlyAsync<ArgumentException>(
            () => store.ReadRowAsync(mismatchedOffset).AsTask());
    }

    [TestMethod]
    public async Task Cell_store_rotates_chunks_and_uses_lru_cache()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
            ChunkSizeBytes = 80,
            RowCacheCapacity = 2,
            RowCacheByteLimit = 1024 * 1024,
        });
        var store = workspace.CreateCellStore();
        var expectedRows = Enumerable.Range(0, 3)
            .Select(index => TestData.Row(
                index,
                (0, TestData.Text(new string((char)('a' + index), 64)))))
            .ToArray();

        foreach (var row in expectedRows)
        {
            await store.AppendRowAsync(row);
        }

        Assert.AreEqual(3, store.ChunkCount);
        Assert.AreEqual(3, store.GetDataFilePaths().Count);
        Assert.IsTrue(store.GetDataFilePaths().All(File.Exists));

        await store.ReadRowAsync(0);
        await store.ReadRowAsync(1);
        await store.ReadRowAsync(0);
        await store.ReadRowAsync(2);

        Assert.AreEqual(2, store.CachedRowCount);
        Assert.IsTrue(store.TryGetCachedRow(0, out _));
        Assert.IsFalse(store.TryGetCachedRow(1, out _));
        Assert.IsTrue(store.TryGetCachedRow(2, out _));

        var actualRows = new List<RowRecord>();
        await foreach (var row in store.ReadRowsAsync(startRowOrdinal: 1, count: 2))
        {
            actualRows.Add(row);
        }

        Assert.AreEqual(2, actualRows.Count);
        TestData.AssertRowsEqual(expectedRows[1], actualRows[0]);
        TestData.AssertRowsEqual(expectedRows[2], actualRows[1]);
    }

    [TestMethod]
    public async Task Failed_append_does_not_publish_an_invalid_row()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var store = workspace.CreateCellStore();
        var invalidRow = new RowRecord(
            0,
            new[]
            {
                new CellRecord(new CellAddress(1, 0), TestData.Text("wrong row")),
            });

        await Assert.ThrowsExactlyAsync<ArgumentException>(
            () => store.AppendRowAsync(invalidRow).AsTask());

        Assert.AreEqual(0L, store.RowCount);
        Assert.AreEqual(16L, store.IndexLengthBytes);
        Assert.AreEqual(16L, store.DataLengthBytes);
    }

    [TestMethod]
    public async Task Store_honors_cancellation_and_rejects_operations_after_dispose()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var store = workspace.CreateCellStore();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsExactlyAsync<OperationCanceledException>(
            () => store.AppendRowAsync(TestData.Row(0), cancellation.Token).AsTask());

        await store.DisposeAsync();
        await Assert.ThrowsExactlyAsync<ObjectDisposedException>(
            () => store.AppendRowAsync(TestData.Row(0)).AsTask());
        Assert.ThrowsExactly<ObjectDisposedException>(() => store.ClearCache());
    }

    [TestMethod]
    public void Text_store_uses_fixed_index_chunks_and_a_bounded_cache()
    {
        using var temporaryDirectory = new TestDirectory();
        var storePath = temporaryDirectory.GetPath("text-store");
        var store = new ChunkedTextStore(storePath, new ChunkedTextStoreOptions
        {
            ChunkSizeBytes = 36,
            MaximumTextSizeBytes = 1024,
            CacheCapacity = 2,
            CacheByteLimit = 1024,
        });
        var values = new[]
        {
            string.Empty,
            "duplicate",
            "duplicate",
            "\u65e5\u672c\u8a9e",
            new string('x', 24),
        };

        for (var index = 0; index < values.Length; index++)
        {
            Assert.AreEqual(index, store.Append(values[index]));
        }
        store.Flush();

        Assert.AreEqual(values.Length, store.Count);
        Assert.IsTrue(store.ChunkCount >= 2);
        Assert.AreEqual(16L + (values.Length * 16L), new FileInfo(store.IndexFilePath).Length);
        for (var index = 0; index < values.Length; index++)
        {
            Assert.AreEqual(values[index], store.Read(index));
        }
        Assert.IsTrue(store.CachedTextCount <= 2);
        Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => store.Read(values.Length));

        store.Dispose();
        Assert.IsFalse(Directory.Exists(storePath));
        Assert.ThrowsExactly<ObjectDisposedException>(() => store.Read(0));
    }

    [TestMethod]
    public async Task Workspace_owns_text_stores_and_cleans_them_on_dispose()
    {
        using var temporaryDirectory = new TestDirectory();
        string workspacePath;
        ChunkedTextStore textStore;
        await using (var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        }))
        {
            workspacePath = workspace.DirectoryPath;
            textStore = workspace.CreateTextStore("shared-strings");
            textStore.Append("value");
            Assert.AreEqual("value", textStore.Read(0));
        }

        Assert.IsFalse(Directory.Exists(workspacePath));
        Assert.ThrowsExactly<ObjectDisposedException>(() => textStore.Read(0));
    }
}
