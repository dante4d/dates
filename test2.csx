#r "nuget: OpenMcdf, 2.4.1"

using System;
using System.IO;
using OpenMcdf;

string filePath = @"C:\Users\dante\Desktop\dates\test\test.xls"; // Update this path

Console.WriteLine("Reading metadata before update...");
ReadMetadata(filePath);

DateTime createdDate = new DateTime(2001, 1, 1, 1, 1, 1);
DateTime modifiedDate = new DateTime(2002, 2, 2, 2, 2, 2);

UpdateXlsMetadata(filePath, createdDate, modifiedDate);

Console.WriteLine("Reading metadata after update...");
ReadMetadata(filePath);

void ReadMetadata(string filePath)
{
    using (var cf = new CompoundFile(filePath, CFSUpdateMode.ReadOnly, CFSConfiguration.Default))
    {
        CFStorage rootStorage = cf.RootStorage;
        if (rootStorage.TryGetStream("\x05SummaryInformation", out CFStream summaryStream))
        {
            byte[] data = summaryStream.GetData();

            DateTime? created = GetProperty(data, 0x0000C);
            DateTime? modified = GetProperty(data, 0x0000D);

            Console.WriteLine($"Content Created: {created}");
            Console.WriteLine($"Date Last Saved: {modified}");
        }
        else
        {
            Console.WriteLine("SummaryInformation stream not found.");
        }
    }
}

void UpdateXlsMetadata(string filePath, DateTime? created, DateTime? modified)
{
    try
    {
        using (var cf = new CompoundFile(filePath, CFSUpdateMode.Update, CFSConfiguration.Default))
        {
            CFStorage rootStorage = cf.RootStorage;

            if (rootStorage.TryGetStream("\x05SummaryInformation", out CFStream summaryStream))
            {
                byte[] data = summaryStream.GetData();

                SetProperty(ref data, 0x0000C, created);  // PIDSI_CREATE_DTM
                SetProperty(ref data, 0x0000D, modified); // PIDSI_LASTSAVE_DTM

                summaryStream.SetData(data);
                cf.Commit();
                Console.WriteLine("Metadata updated successfully.");
            }
            else
            {
                Console.WriteLine("SummaryInformation stream not found.");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error updating metadata: {ex.Message}");
    }
}

DateTime? GetProperty(byte[] data, int propertyId)
{
    int offset = FindPropertyOffset(data, propertyId);
    if (offset >= 0)
    {
        long fileTime = BitConverter.ToInt64(data, offset);
        return DateTime.FromFileTime(fileTime);
    }
    return null;
}

void SetProperty(ref byte[] data, int propertyId, DateTime? value)
{
    if (!value.HasValue) return;

    int offset = FindPropertyOffset(data, propertyId);
    if (offset >= 0)
    {
        byte[] timeBytes = BitConverter.GetBytes(value.Value.ToFileTime());
        Buffer.BlockCopy(timeBytes, 0, data, offset, 8);
    }
}

int FindPropertyOffset(byte[] data, int propertyId)
{
    for (int i = 0; i < data.Length - 4; i += 4)
    {
        int id = BitConverter.ToInt32(data, i);
        if (id == propertyId)
        {
            return i + 8; // Offset where the DateTime value starts
        }
    }
    return -1;
}
