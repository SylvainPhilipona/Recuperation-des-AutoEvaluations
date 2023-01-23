function Replace-WordContent{



























    return






    param(
        [Parameter(Mandatory=$true)]
        [string]$Template,

        [Parameter(Mandatory=$true)]
        [hashtable]$ContentToReplace,

        [Parameter(Mandatory=$true)]
        [string]$NewPath,

        [Parameter(Mandatory=$true)]
        [string]$NewName,

        [string]$tempFolder = ($env:TEMP + "\Populate-Word-DOCX")
    )

    # unzip function
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    function Unzip {
        param([string]$zipfile, [string]$outpath)
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    }
    function Zip {
        param([string]$folderInclude, [string]$outZip)
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


    # prepare folder
    Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
    mkdir $tempFolder | Out-Null

    # unzip DOCX
    Unzip $template $tempFolder

    # replace text
    $bodyFile = $tempFolder + "\word\document.xml"
    $body = Get-Content $bodyFile -Encoding UTF8

    foreach ($i in $ContentToReplace.GetEnumerator()) {
        $body = $body.Replace("$($i.Name)", "$($i.Value)")
    }
    $body | Out-File $bodyFile -Force -Encoding utf8

    # zip DOCX
    if(!(Test-Path -Path $NewPath)){mkdir $NewPath | Out-Null}
    $destfile = "$NewPath\$NewName.docx"
    Remove-Item $destfile -Force -ErrorAction SilentlyContinue
    Zip $tempFolder $destfile

    # Convert to PDF
    $word_app = New-Object -ComObject Word.Application
    $document = $word_app.Documents.Open($destfile)

    $pdf_filename = "$NewPath\$NewName.pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)

    $document.Close()
    $word_app.Quit()

    # clean temp folder
    Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null

    # Delete the word file
    Remove-Item $destfile
}