$File = "C:\Users\dante\Desktop\dates\test\test.xls"
$Shell = New-Object -ComObject Shell.Application
$Folder = $Shell.Namespace((Get-Item $File).DirectoryName)
$FileItem = $Folder.ParseName((Get-Item $File).Name)

# Iterate through all known metadata fields (0 to 287 is the typical range)
for ($i = 0; $i -lt 288; $i++) {
    $Key = $Folder.GetDetailsOf($Folder.Items, $i)
    $Value = $Folder.GetDetailsOf($FileItem, $i)
    
    if ($Key -and $Value) {
        Write-Output "$Key : $Value"
    }
}
