#r "nuget: Spectre.Console, 0.49.2-preview.0.69"
#r "nuget: DocumentFormat.OpenXml, 3.2.0"
#r "nuget: ClosedXML, 0.104.2"
#r "nuget: NPOI, 2.7.2"
#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Spectre.Console;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;

var log = true;

var sheetNum = 1;
var skipCount = 3;

record Entry {
    public required string path { get; init; }
    public DateTime? oldCreationTime   { get; init; }
    public DateTime? oldLastWriteTime  { get; init; }
    public DateTime? oldLastAccessTime { get; init; }
    public DateTime? oldPropsCreated   { get; init; }
    public DateTime? oldPropsModified  { get; init; }
    public DateTime? newCreationTime   { get; init; }
    public DateTime? newLastWriteTime  { get; init; }
    public DateTime? newLastAccessTime { get; init; }
    public DateTime? newPropsCreated   { get; init; }
    public DateTime? newPropsModified  { get; init; }    
}

DateTime? getDateOnly(IXLRow row, int index) =>
    row.Cell(index).TryGetValue<DateTime>(out var value) ? value.Date : (DateTime?)null;

List<Entry> loadEntries(string folder, string list) =>
    new XLWorkbook(Path.Combine(folder, list))
        .Worksheet(sheetNum)
        .RowsUsed()
        .Skip(skipCount)
        .Select(row => new Entry {
            path              = Path.Combine(folder, row.Cell(1).GetString()),
            // old
            oldPropsCreated   = getDateOnly(row, 2),
            oldPropsModified  = getDateOnly(row, 3),
            oldCreationTime   = getDateOnly(row, 4),
            oldLastWriteTime  = getDateOnly(row, 5),
            oldLastAccessTime = getDateOnly(row, 6),
            // new
            newPropsCreated   = getDateOnly(row, 7),
            newPropsModified  = getDateOnly(row, 8),
            newCreationTime   = getDateOnly(row, 9),
            newLastWriteTime  = getDateOnly(row, 10),
            newLastAccessTime = getDateOnly(row, 11)
        })
        .ToList();

void saveEntries(string folder, string list, List<Entry> entries) {
    string path = Path.Combine(folder, list);

    string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
    string name = $"{Path.GetFileNameWithoutExtension(path)}-{timestamp}{Path.GetExtension(path)}";
    string newPath = Path.Combine(Path.GetDirectoryName(path) ?? ".", name);

    File.Copy(path, newPath);

    using var workbook = new XLWorkbook(newPath);
    var worksheet = workbook.Worksheet(sheetNum);

    int index = skipCount + 1;
    foreach (var entry in entries) {
        var row = worksheet.Row(index++);

        row.Cell(1).SetValue(entry.path);
        // old
        row.Cell(2).SetValue(entry.oldPropsCreated?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(3).SetValue(entry.oldPropsModified?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(4).SetValue(entry.oldCreationTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(5).SetValue(entry.oldLastWriteTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(6).SetValue(entry.oldLastAccessTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        // new
        row.Cell(7).SetValue(entry.newPropsCreated?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(8).SetValue(entry.newPropsModified?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(9).SetValue(entry.newCreationTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(10).SetValue(entry.newLastWriteTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
        row.Cell(11).SetValue(entry.newLastAccessTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");
    }

    workbook.Save();

    Console.WriteLine($"Výsledek uložen do {newPath}");
}

DateTime? replaceDate(DateTime? oldDateTime, DateTime? newDateTime) {
    if (oldDateTime.HasValue && newDateTime.HasValue) {
        return new DateTime(
            newDateTime.Value.Year,
            newDateTime.Value.Month,
            newDateTime.Value.Day,
            oldDateTime.Value.Hour,
            oldDateTime.Value.Minute,
            oldDateTime.Value.Second
        );
    }
    return null;
}

Entry processEntry(Entry entry, OpenXmlPackage? doc = null) {    
    if (log) AnsiConsole.WriteLine(entry.ToString());

    var props = doc?.PackageProperties;

    if (doc == null) AnsiConsole.MarkupLine($"[yellow]doc == null[/]");
    if (props == null) AnsiConsole.MarkupLine($"[yellow]props == null[/]");
    if (props?.Created == null) AnsiConsole.MarkupLine($"[yellow]props.Created == null[/]");
    if (props?.Modified == null) AnsiConsole.MarkupLine($"[yellow]props.Modified == null[/]");

    var oldCreationTime = File.GetCreationTime(entry.path);
    var oldLastWriteTime = File.GetLastWriteTime(entry.path);
    var oldLastAccessTime = File.GetLastAccessTime(entry.path);

    var newEntry = entry with {
        path = entry.path,
        oldCreationTime   = oldCreationTime,
        oldLastWriteTime  = oldLastWriteTime,
        oldLastAccessTime = oldLastAccessTime,
        oldPropsCreated   = props?.Created,
        oldPropsModified  = props?.Modified,
        newCreationTime   = replaceDate(oldCreationTime, entry.newCreationTime),
        newLastWriteTime  = replaceDate(oldLastWriteTime, entry.newLastWriteTime),
        newLastAccessTime = replaceDate(oldLastAccessTime, entry.newLastAccessTime),
        newPropsCreated   = replaceDate(props?.Created, entry.newPropsCreated),
        newPropsModified  = replaceDate(props?.Modified, entry.newPropsModified)
    };

    if (log) AnsiConsole.WriteLine(newEntry.ToString());

    if (props != null) {
        if (newEntry.newPropsCreated.HasValue) props.Created = newEntry.newPropsCreated;
        if (newEntry.newPropsModified.HasValue) props.Modified = newEntry.newPropsModified;
        doc!.Dispose();
    }

    using (var stream = new FileStream(newEntry.path, FileMode.Open, FileAccess.Write, FileShare.None)) {
        stream.Flush(true);
    }    

    if (newEntry.newCreationTime.HasValue) {
        File.SetCreationTime(newEntry.path, newEntry.newCreationTime.Value);
    }
    if (newEntry.newLastWriteTime.HasValue) {
        // if (log) AnsiConsole.MarkupLine($"[yellow]LastWriteTime == {newEntry.newLastWriteTime}[/]");
        File.SetLastWriteTime(newEntry.path, newEntry.newLastWriteTime.Value);
    }
    if (newEntry.newLastAccessTime.HasValue) {
        // if (log) AnsiConsole.MarkupLine($"[yellow]LastAccessTime == {newEntry.newLastAccessTime}[/]");
        File.SetLastAccessTime(newEntry.path, newEntry.newLastAccessTime.Value);
    }        

    return newEntry;
}

Entry processXlsEntry(Entry entry) {    
    if (log) AnsiConsole.MarkupLine($"[green]Processing file: {entry.path}[/]");

    if (!File.Exists(entry.path))
    {
        throw new Exception($"File not found: {entry.path}");
    }

    try
    {
        byte[] fileBytes = File.ReadAllBytes(entry.path); // ✅ Load file into memory
        using var memStream = new MemoryStream(fileBytes);
        
        var system = new POIFSFileSystem(memStream); // ❌ No `using` here!

        SummaryInformation? information = null;

        try
        {
            AnsiConsole.MarkupLine("[blue]Checking for metadata entry...[/]");

            if (system.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME))
            {
                var entry2 = system.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                if (entry2 is DocumentNode node)
                {
                    using var input = new DocumentInputStream(node);
                    information = (SummaryInformation) PropertySetFactory.Create(input);
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]Metadata not found, creating new one.[/]");
                information = PropertySetFactory.CreateSummaryInformation();
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error reading metadata: {ex.Message}[/]");
            AnsiConsole.MarkupLine($"[red]Stack Trace: {ex.StackTrace}[/]");
        }

        // if (information != null)
        // {
        //     try
        //     {
        //         AnsiConsole.MarkupLine("[blue]Updating metadata...[/]");
        //         if (entry.newPropsCreated.HasValue) information.CreateDateTime = entry.newPropsCreated.Value;
        //         if (entry.newPropsModified.HasValue) information.LastSaveDateTime = entry.newPropsModified.Value;

        //         using var output = new MemoryStream();
        //         information.Write(output);
        //         var bytes = output.ToArray();

        //         if (system.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME))
        //         {
        //             var existingEntry = system.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
        //             if (existingEntry is EntryNode entryToRemove)
        //             {
        //                 system.Root.DeleteEntry(entryToRemove);
        //             }
        //         }

        //         system.Root.CreateDocument(SummaryInformation.DEFAULT_STREAM_NAME, new MemoryStream(bytes, writable: false));
        //     }
        //     catch (Exception ex)
        //     {
        //         AnsiConsole.MarkupLine($"[red]Error updating metadata: {ex.Message}[/]");
        //         AnsiConsole.MarkupLine($"[red]Stack Trace: {ex.StackTrace}[/]");
        //     }
        // }

        if (information != null)
        {
            try
            {
                AnsiConsole.MarkupLine($"[blue]Old Created: {information.CreateDateTime}[/]");
                AnsiConsole.MarkupLine($"[blue]Old Modified: {information.LastSaveDateTime}[/]");

                if (entry.newPropsCreated.HasValue) information.CreateDateTime = entry.newPropsCreated.Value;
                if (entry.newPropsModified.HasValue) information.LastSaveDateTime = entry.newPropsModified.Value;

                AnsiConsole.MarkupLine($"[green]New Created: {information.CreateDateTime}[/]");
                AnsiConsole.MarkupLine($"[green]New Modified: {information.LastSaveDateTime}[/]");

                // using var output = new MemoryStream();
                // information.Write(output);
                // var bytes = output.ToArray();

                // if (system.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME))
                // {
                //     var existingEntry = system.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                //     if (existingEntry is EntryNode entryToRemove)
                //     {
                //         system.Root.DeleteEntry(entryToRemove);
                //     }
                // }

                // system.Root.CreateDocument(SummaryInformation.DEFAULT_STREAM_NAME, new MemoryStream(bytes, writable: false));

                using var output = new MemoryStream();
                information.Write(output);
                var bytes = output.ToArray();

                if (system.Root.HasEntry(SummaryInformation.DEFAULT_STREAM_NAME))
                {
                    var existingEntry = system.Root.GetEntry(SummaryInformation.DEFAULT_STREAM_NAME);
                    if (existingEntry is EntryNode entryToRemove)
                    {
                        system.Root.DeleteEntry(entryToRemove);
                    }
                }

                system.Root.CreateDocument(SummaryInformation.DEFAULT_STREAM_NAME, new MemoryStream(bytes, writable: false));

                // ✅ Force write system update
                system.WriteFileSystem(memStream);                
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error updating metadata: {ex.Message}[/]");
                AnsiConsole.MarkupLine($"[red]Stack Trace: {ex.StackTrace}[/]");
            }
        }

        try
        {
            AnsiConsole.MarkupLine("[blue]Writing memory copy back to disk...[/]");
            File.WriteAllBytes(entry.path, memStream.ToArray());
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error writing file: {ex.Message}[/]");
            AnsiConsole.MarkupLine($"[red]Stack Trace: {ex.StackTrace}[/]");
        }

        system.Close(); // ✅ Explicitly close POIFSFileSystem
    }
    catch (Exception ex)
    {
        AnsiConsole.MarkupLine($"[red]Fatal error: {ex.Message}[/]");
        AnsiConsole.MarkupLine($"[red]Stack Trace: {ex.StackTrace}[/]");
    }

    return entry;
}

Func<Entry, Entry> getHandler(string extension) =>
    extension.ToLower() switch {
        ".docx" => (entry) => processEntry(entry, WordprocessingDocument.Open(entry.path, true)),
        ".xlsx" => (entry) => processEntry(entry, SpreadsheetDocument.Open(entry.path, true)),
        // ".xls"  => (entry) => processXlsEntry(entry),
        _ => (entry) => processEntry(entry)
    };

Entry process(Entry entry) {
    // try {
        AnsiConsole.MarkupLine($"[blue]Zpracovávám {entry.path}...[/]");
        var extension = Path.GetExtension(entry.path);

        var handler = getHandler(extension);
        var newEntry = handler(entry);

        return newEntry;
    // } catch (Exception e) {
    //     AnsiConsole.MarkupLine($"[red]Chyba: {e.Message}[/]");
    //     return entry;
    // }
}

var args = Environment.GetCommandLineArgs();

string folder = args.Length > 2 ? args[2] : "test";

if (!Directory.Exists(folder)) {
    AnsiConsole.MarkupLine($"[red]Složka {folder} neexistuje.[/]");
    return;
}

AnsiConsole.MarkupLine($"[blue]Zpracování složky {folder}...[/]");

var entries = loadEntries(folder, "list.xlsx");
var newEntries = entries.Select(process).ToList();
saveEntries(folder, "list.xlsx", newEntries);
AnsiConsole.MarkupLine("[green]Soubory zpracovány![/]");
