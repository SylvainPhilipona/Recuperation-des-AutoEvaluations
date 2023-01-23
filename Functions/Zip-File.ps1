function Zip-File{
    param([string]$outZip, [string]$folderInclude)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.CompressionLevel]$compression = "Optimal"
    $ziparchive = [System.IO.Compression.ZipFile]::Open( $outZip, "Update" )

    # loop all child files
    $realtiveTempFolder = (Resolve-Path $tempFolder -Relative).TrimStart(".\")
    foreach ($file in (Get-ChildItem $folderInclude -Recurse)) {
        # skip directories
        if ($file.GetType().ToString() -ne "System.IO.DirectoryInfo") {
            # relative path
            $relpath = ""
            if ($file.FullName) {
                $relpath = (Resolve-Path $file.FullName -Relative)
            }
            if (!$relpath) {
                $relpath = $file.Name
            } else {
                $relpath = $relpath.Replace($realtiveTempFolder, "")
                $relpath = $relpath.TrimStart(".\").TrimStart("\\")
            }

            # display
            Write-Host $relpath -Fore Green
            Write-Host $file.FullName -Fore Yellow

            # add file
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ziparchive, $file.FullName, $relpath, $compression) | Out-Null
        }
    }
    $ziparchive.Dispose()
}