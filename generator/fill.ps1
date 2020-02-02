# global vars
$template = "C:\Users\adrie\OneDrive\Documents\COOP\cover_letters_auto\master_coverletter_template.docx"
$tempFolder = "C:\Users\adrie\OneDrive\Documents\COOP\cover_letters_auto\TEMP"

# unzip function
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    param([string]$zipfile, [string]$outpath)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

# zip function
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

            # add file
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ziparchive, $file.FullName, $relpath, $compression) | Out-Null
        }
    }
    $ziparchive.Dispose()
}

# insert function
function Replace {
    param([string]$keyword, [string]$value)
    $final = "[" + $keyword + "]"
    
    return $body.Replace($final, $value)
}

# prepare folder
Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null
mkdir $tempFolder | Out-Null

# unzip DOCX
Unzip $template $tempFolder

# replace text
$bodyFile = $tempFolder + "\word\document.xml"
$body = Get-Content $bodyFile
$body = Replace "name" "bob marley"
Write-Host $body -Fore Green
$body | Out-File $bodyFile -Force -Encoding ascii

# zip DOCX
$destfile = $template.Replace(".docx", "-after.docx")
Remove-Item $destfile -Force -ErrorAction SilentlyContinue
Zip $tempFolder $destfile

# remove temp folder
Remove-Item $tempFolder -ErrorAction SilentlyContinue -Recurse -Confirm:$false | Out-Null