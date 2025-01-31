#nullable enable
#r "nuget: Spectre.Console, 0.49.2-preview.0.69"
#r "nuget: DocumentFormat.OpenXml, 3.2.0"
#r "nuget: ClosedXML, 0.104.2"

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Spectre.Console;

var log = true;
const int SheetNum = 1;
const int SkipCount = 3;

record Entry(
    string Path,
    DateTime? OldCreationTime = null,
    DateTime? OldLastWriteTime = null,
    DateTime? OldLastAccessTime = null,
    DateTime? OldPropsCreated = null,
    DateTime? OldPropsModified = null,
    DateTime? NewCreationTime = null,
    DateTime? NewLastWriteTime = null,
    DateTime? NewLastAccessTime = null,
    DateTime? NewPropsCreated = null,
    DateTime? NewPropsModified = null
);

static class Extensions
{
    public static DateTime? GetDateOnly(this IXLRow row, int index) =>
        row.Cell(index).TryGetValue<DateTime>(out var value) ? value.Date : null;

    public static void SetCellDate(this IXLCell cell, DateTime? date) =>
        cell.SetValue(date?.ToString("yyyy-MM-dd HH:mm:ss") ?? "");

    public static void LogIf(this bool condition, string message) {
        if (condition) AnsiConsole.WriteLine(message);
    }
}

List<Entry> LoadEntries(string filePath) =>
    new XLWorkbook(filePath)
        .Worksheet(SheetNum)
        .RowsUsed()
        .Skip(SkipCount)
        .Select(row => new Entry(
            row.Cell(1).GetString(),
            OldPropsCreated: row.GetDateOnly(2),
            OldPropsModified: row.GetDateOnly(3),
            OldCreationTime: row.GetDateOnly(4),
            OldLastWriteTime: row.GetDateOnly(5),
            OldLastAccessTime: row.GetDateOnly(6),
            NewPropsCreated: row.GetDateOnly(7),
            NewPropsModified: row.GetDateOnly(8),
            NewCreationTime: row.GetDateOnly(9),
            NewLastWriteTime: row.GetDateOnly(10),
            NewLastAccessTime: row.GetDateOnly(11)
        ))
        .ToList();

void SaveEntries(string filePath, List<Entry> entries)
{
    string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
    string newPath = Path.Combine(Path.GetDirectoryName(filePath) ?? ".", 
                    $"{Path.GetFileNameWithoutExtension(filePath)}-{timestamp}{Path.GetExtension(filePath)}");

    File.Copy(filePath, newPath);

    using var workbook = new XLWorkbook(newPath);
    var worksheet = workbook.Worksheet(SheetNum);

    for (int i = 0; i < entries.Count; i++)
    {
        var row = worksheet.Row(SkipCount + 1 + i);
        SetRowCells(row, entries[i]);
    }

    workbook.Save();
    AnsiConsole.WriteLine($"Saved results to {newPath}");
}

void SetRowCells(IXLRow row, Entry entry)
{
    row.Cell(1).SetValue(entry.Path);
    row.Cell(2).SetCellDate(entry.OldPropsCreated);
    row.Cell(3).SetCellDate(entry.OldPropsModified);
    row.Cell(4).SetCellDate(entry.OldCreationTime);
    row.Cell(5).SetCellDate(entry.OldLastWriteTime);
    row.Cell(6).SetCellDate(entry.OldLastAccessTime);
    row.Cell(7).SetCellDate(entry.NewPropsCreated);
    row.Cell(8).SetCellDate(entry.NewPropsModified);
    row.Cell(9).SetCellDate(entry.NewCreationTime);
    row.Cell(10).SetCellDate(entry.NewLastWriteTime);
    row.Cell(11).SetCellDate(entry.NewLastAccessTime);
}

Entry ProcessEntry(Entry entry, OpenXmlPackage? doc = null)
{
    log.LogIf(entry.ToString());

    var props = doc?.PackageProperties;

    DateTime? ReplaceDate(DateTime? oldDate, DateTime? newDate) =>
        (oldDate.HasValue && newDate.HasValue)
        ? new DateTime(newDate.Value.Year, newDate.Value.Month, newDate.Value.Day,
                       oldDate.Value.Hour, oldDate.Value.Minute, oldDate.Value.Second)
        : null;

    var updatedEntry = entry with
    {
        OldCreationTime = File.GetCreationTime(entry.Path),
        OldLastWriteTime = File.GetLastWriteTime(entry.Path),
        OldLastAccessTime = File.GetLastAccessTime(entry.Path),
        OldPropsCreated = props?.Created,
        OldPropsModified = props?.Modified,
        NewCreationTime = ReplaceDate(entry.OldCreationTime, entry.NewCreationTime),
        NewLastWriteTime = ReplaceDate(entry.OldLastWriteTime, entry.NewLastWriteTime),
        NewLastAccessTime = ReplaceDate(entry.OldLastAccessTime, entry.NewLastAccessTime),
        NewPropsCreated = ReplaceDate(props?.Created, entry.NewPropsCreated),
        NewPropsModified = ReplaceDate(props?.Modified, entry.NewPropsModified)
    };

    ApplyEntryUpdates(updatedEntry, props, doc);

    return updatedEntry;
}

void ApplyEntryUpdates(Entry entry, OpenXmlPackageProperties? props, OpenXmlPackage? doc)
{
    props?.Apply(p =>
    {
        if (entry.NewPropsCreated.HasValue) p.Created = entry.NewPropsCreated;
        if (entry.NewPropsModified.HasValue) p.Modified = entry.NewPropsModified;
    });

    doc?.Dispose();

    if (entry.NewCreationTime.HasValue) File.SetCreationTime(entry.Path, entry.NewCreationTime.Value);
    if (entry.NewLastWriteTime.HasValue) File.SetLastWriteTime(entry.Path, entry.NewLastWriteTime.Value);
    if (entry.NewLastAccessTime.HasValue) File.SetLastAccessTime(entry.Path, entry.NewLastAccessTime.Value);
}

Func<Entry, Entry> GetHandler(string extension) => extension.ToLower() switch
{
    ".docx" => entry => ProcessEntry(entry, WordprocessingDocument.Open(entry.Path, true)),
    ".xlsx" => entry => ProcessEntry(entry, SpreadsheetDocument.Open(entry.Path, true)),
    _ => ProcessEntry
};

List<Entry> ProcessEntries(List<Entry> entries) =>
    entries.Select(entry =>
    {
        log.LogIf($"Processing {entry.Path}...");
        try
        {
            var handler = GetHandler(Path.GetExtension(entry.Path));
            return handler(entry);
        }
        catch (Exception e)
        {
            AnsiConsole.MarkupLine($"[red]Error: {e.Message}[/]");
            return entry;
        }
    }).ToList();

var entries = LoadEntries("list.xlsx");
var processedEntries = ProcessEntries(entries);
SaveEntries("list.xlsx", processedEntries);

AnsiConsole.MarkupLine("[green]Files processed![/]");
