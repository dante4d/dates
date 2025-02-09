#r "nuget: Spectre.Console, 0.49.2-preview.0.69"
#r "nuget: ClosedXML, 0.104.2"
#r "nuget: DocumentFormat.OpenXml, 3.2.0"
#r "nuget: NPOI, 2.7.2"
#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Threading;
using Spectre.Console;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;

record Stamps {
    public DateTime? CreationTime { get; init; }
    public DateTime? LastWriteTime { get; init; }
    public DateTime? LastAccessTime { get; init; }
    public DateTime? PropsCreated { get; init; }
    public DateTime? PropsModified { get; init; }
}

record Entry {
    public required string Path { get; init; }
    public Stamps? Old { get; init; }
    public Stamps? New { get; init; }
}

interface IHandler {
    Stamps Read(string path);
    void Write(string path, Stamps stamps);
}

class DefaultHandler : IHandler {
    public virtual Stamps Read(string file) =>
        new() {
            CreationTime = File.GetCreationTime(file),
            LastWriteTime = File.GetLastWriteTime(file),
            LastAccessTime = File.GetLastAccessTime(file)
        };

    public virtual void Write(string file, Stamps stamps) {
        if (stamps.CreationTime.HasValue) File.SetCreationTime(file, stamps.CreationTime.Value);
        if (stamps.LastWriteTime.HasValue) File.SetLastWriteTime(file, stamps.LastWriteTime.Value);
        if (stamps.LastAccessTime.HasValue) File.SetLastAccessTime(file, stamps.LastAccessTime.Value);
    }
}

class XmlHandler : DefaultHandler {
    private OpenXmlPackage Open(string file) =>
        Path.GetExtension(file).ToLower() switch {
            ".docx" => WordprocessingDocument.Open(file, true),
            ".xlsx" => SpreadsheetDocument.Open(file, true),
            _ => throw new NotSupportedException($"Unsupported file type: {file}")
        };

    public override Stamps Read(string file) {
        try{
            using var document = Open(file);
            var properties = document.PackageProperties;
            return base.Read(file) with {
                PropsCreated = properties.Created,
                PropsModified = properties.Modified
            };
        } catch (Exception ex) {
            AnsiConsole.MarkupLine($"[red]Error reading XML metadata: {ex.Message}[/]");
        }
        return base.Read(file);
    }

    public override void Write(string file, Stamps stamps) {
        try {
            using var document = Open(file);
            var properties = document.PackageProperties;
            if (stamps.PropsCreated.HasValue) properties.Created = stamps.PropsCreated;
            if (stamps.PropsModified.HasValue) properties.Modified = stamps.PropsModified;
        } catch (Exception ex) {
            AnsiConsole.MarkupLine($"[red]Error writing XML metadata: {ex.Message}[/]");
        } finally {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            base.Write(file, stamps);
        }
    }
}

class OleHandler : DefaultHandler {
    public override Stamps Read(string file) {
        try {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            var poifs = new POIFSFileSystem(fs);

            if (!poifs.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME)) {
                AnsiConsole.MarkupLine($"[red]No SummaryInformation stream found in {file}[/]");
            } else {
                var summaryEntry = poifs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                if (summaryEntry is not DocumentEntry documentEntry) {
                    AnsiConsole.MarkupLine($"[red]SummaryInformation stream is not a document entry in {file}[/]");
                } else {
                    using var input = new DocumentInputStream(documentEntry);
                    var summaryInfo = (SummaryInformation)PropertySetFactory.Create(input);

                    return base.Read(file) with {
                        PropsCreated = summaryInfo.CreateDateTime,
                        PropsModified = summaryInfo.LastSaveDateTime
                    };
                }
            }
        } catch (Exception ex) {
            AnsiConsole.MarkupLine($"[red]Error reading OLE metadata: {ex.Message}[/]");
        }
        return base.Read(file);
    }

    public override void Write(string file, Stamps stamps) {
        try {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite);
            var poifs = new POIFSFileSystem(fs);

            if (!poifs.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME)) {
                AnsiConsole.MarkupLine($"[red]No SummaryInformation stream found in {file}[/]");
            } else {
                var summaryEntry = poifs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                if (summaryEntry is not DocumentEntry documentEntry) {
                    AnsiConsole.MarkupLine($"[red]SummaryInformation stream is not a document entry in {file}[/]");
                } else {
                    using var input = new DocumentInputStream(documentEntry);
                    var summaryInfo = (SummaryInformation)PropertySetFactory.Create(input);

                    if (stamps.PropsCreated.HasValue) summaryInfo.CreateDateTime = stamps.PropsCreated.Value;
                    if (stamps.PropsModified.HasValue) summaryInfo.LastSaveDateTime = stamps.PropsModified.Value;

                    using var output = new MemoryStream();
                    summaryInfo.Write(output);
                    var bytes = output.ToArray();

                    if (poifs.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME) is EntryNode entryToRemove) {
                        poifs.Root.DeleteEntry(entryToRemove);
                    }

                    poifs.Root.CreateDocument(SummaryInformation.DEFAULT_STREAM_NAME, new MemoryStream(bytes, false));

                    using var outFs = new FileStream(file, FileMode.Create, FileAccess.Write);
                    poifs.WriteFileSystem(outFs);
                }
            }
        } catch (Exception ex) {
            AnsiConsole.MarkupLine($"[red]Error writing OLE metadata: {ex.Message}[/]");
        } finally {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            base.Write(file, stamps);
        }
    }
}

class Processor {
    static readonly string List = "list.xlsx";
    static readonly int SheetNum = 1;
    static readonly int SkipLines = 3;

    static readonly Dictionary<string, IHandler> Handlers = new() {
        { ".docx", new XmlHandler() },
        { ".xlsx", new XmlHandler() },
        { ".xls", new OleHandler() }
    };

    static readonly IHandler DefaultHandler = new DefaultHandler();

    public static void Run(string[] args) {
        var folder = args.Length > 2 ? args[2] : "test";
        if (Directory.Exists(folder)) {
            AnsiConsole.MarkupLine($"[blue]Processing folder {folder}...[/]");
        } else {
            AnsiConsole.MarkupLine($"[red]Folder {folder} does not exist.[/]");
            return;
        }

        var file = Path.Combine(folder, List);

        // load entries from file
        AnsiConsole.MarkupLine($"[blue]Loading entries from {file}...[/]");
        var entries = LoadEntries(file);

        // read old timestamps
        AnsiConsole.MarkupLine($"[blue]Reading old timestamps...[/]");
        entries = ReadStamps(folder, entries);

        // update timestamps (time from old to new)
        AnsiConsole.MarkupLine($"[blue]Updating new timestamps...[/]");
        entries = entries.Select(UpdateStamps).ToList();

        // write new timestamps
        AnsiConsole.MarkupLine($"[blue]Writing new timestamps...[/]");
        WriteStamps(folder, entries);

        // copy entries file
        AnsiConsole.MarkupLine($"[blue]Copying entries file...[/]");
        var newFile = CopyEntries(file);

        // save entries to file
        AnsiConsole.MarkupLine($"[blue]Saving entries to {newFile}...[/]");
        SaveEntries(newFile, entries);

        AnsiConsole.MarkupLine($"[green]Folder {folder} processed![/]");
    }

    static IHandler GetHandler(string file) =>
        Handlers.TryGetValue(Path.GetExtension(file).ToLower(), out var handler) ? handler : DefaultHandler;

    static List<Entry> ReadStamps(string folder, List<Entry> entries) =>
        entries.Select(entry => {
            var handler = GetHandler(entry.Path);
            var stamps = handler.Read(Path.Combine(folder, entry.Path));
            return entry with { Old = stamps };
        }).ToList();

    static DateTime? ReplaceDate(DateTime? oldDateTime, DateTime? newDateTime) {
        if (oldDateTime is null || newDateTime is null) {
            return null;
        } else {
            return new DateTime(
                newDateTime.Value.Year,
                newDateTime.Value.Month,
                newDateTime.Value.Day,
                oldDateTime.Value.Hour,
                oldDateTime.Value.Minute,
                oldDateTime.Value.Second
            );
        }

    }

    static Entry UpdateStamps(Entry entry) {
        if (entry.Old is null || entry.New is null) {
            return entry;
        } else {
            return entry with {
                New = new () {
                    CreationTime = ReplaceDate(entry.Old.CreationTime, entry.New.CreationTime),
                    LastWriteTime = ReplaceDate(entry.Old.LastWriteTime, entry.New.LastWriteTime),
                    LastAccessTime = ReplaceDate(entry.Old.LastAccessTime, entry.New.LastAccessTime),
                    PropsCreated = ReplaceDate(entry.Old.PropsCreated, entry.New.PropsCreated),
                    PropsModified = ReplaceDate(entry.Old.PropsModified, entry.New.PropsModified)
                }
            };
        }
    }

    static void WriteStamps(string folder, List<Entry> entries) {
        foreach (var entry in entries) {
            if (entry.New is null) continue;
            var handler = GetHandler(entry.Path);
            handler.Write(Path.Combine(folder, entry.Path), entry.New);
        }
    }

    static DateTime? GetDateTime(IXLRow row, int index) =>
        row.Cell(index).TryGetValue<DateTime>(out var value) ? value : (DateTime?)null;

    static List<Entry> LoadEntries(string file) =>
        new XLWorkbook(file)
            .Worksheet(SheetNum)
            .RowsUsed()
            .Skip(SkipLines)
            .Select(row => new Entry {
                Path = row.Cell(1).GetString(),
                Old = new() {
                    CreationTime = GetDateTime(row, 2),
                    LastWriteTime = GetDateTime(row, 3),
                    LastAccessTime = GetDateTime(row, 4),
                    PropsCreated = GetDateTime(row, 5),
                    PropsModified = GetDateTime(row, 6)
                },
                New = new() {
                    CreationTime = GetDateTime(row, 7),
                    LastWriteTime = GetDateTime(row, 8),
                    LastAccessTime = GetDateTime(row, 9),
                    PropsCreated = GetDateTime(row, 10),
                    PropsModified = GetDateTime(row, 11)
                }
            })
            .ToList();

    static string CopyEntries(string file) {
        var name = Path.GetFileNameWithoutExtension(file);
        var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var extension = Path.GetExtension(file);
        var path1 = Path.GetDirectoryName(file) ?? ".";
        var path2 = $"{name}-{stamp}{extension}";
        var newFile = Path.Combine(path1, path2);
        File.Copy(file, newFile);
        return newFile;
    }

    static void SaveEntries(string file, List<Entry> entries) {
        using (var workbook = new XLWorkbook(file)) {
            var worksheet = workbook.Worksheet(SheetNum);

            var index = SkipLines + 1;
            foreach (var entry in entries) {
                var row = worksheet.Row(index++);

                row.Cell(1).SetValue(entry.Path);
                // old
                row.Cell(2).SetValue(entry.Old?.PropsCreated?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(3).SetValue(entry.Old?.PropsModified?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(4).SetValue(entry.Old?.CreationTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(5).SetValue(entry.Old?.LastWriteTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(6).SetValue(entry.Old?.LastAccessTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                // new
                row.Cell(7).SetValue(entry.New?.PropsCreated?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(8).SetValue(entry.New?.PropsModified?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(9).SetValue(entry.New?.CreationTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(10).SetValue(entry.New?.LastWriteTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
                row.Cell(11).SetValue(entry.New?.LastAccessTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
            }

            workbook.Save();
        }
    }
}

Processor.Run(Environment.GetCommandLineArgs())
