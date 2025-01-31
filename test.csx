#r "nuget: DocumentFormat.OpenXml, 3.2.0"

using System;
using System.IO;

using DocumentFormat.OpenXml.Packaging;

string path = @"C:\Users\dante\Desktop\dates\test\test.docx";
// string path = @"C:\Users\dante\Desktop\dates\test\test.xlsx";

var time = DateTime.Parse("2025-01-01T06:06:06Z");

using (var document = WordprocessingDocument.Open(path, true))
// using (var document = SpreadsheetDocument.Open(path, true))
{
	// Získání metadat dokumentu
	var props = document.PackageProperties;

	// Nastavení nového data vytvoření a uložení
	props.Created = time;
	props.Modified = time;

	Console.WriteLine("Metadata byla změněna.");
}

// Nastav nová data
DateTime creationTime = new DateTime(2025, 1, 1, 6, 6, 6);
DateTime lastWriteTime = new DateTime(2025, 1, 1, 6, 6, 6);
DateTime lastAccessTime = new DateTime(2025, 1, 1, 6, 6, 6);

// Změň atributy
File.SetCreationTime(path, creationTime);
File.SetLastWriteTime(path, lastWriteTime);
File.SetLastAccessTime(path, lastAccessTime);

// Ověření změn
Console.WriteLine("Creation Time: " + File.GetCreationTime(path));
Console.WriteLine("Last Write Time: " + File.GetLastWriteTime(path));
Console.WriteLine("Last Access Time: " + File.GetLastAccessTime(path));
