# Modified Function from https://gallery.technet.microsoft.com/office/Insert-pictures-in-a-dad235ac 
Function Add-OSCPicture {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('wordpath')]
        [String]$WordDocumentPath,
        [Parameter(Mandatory=$true,Position=1)]
        [Alias('imgpath')]
        [String]$ImageFolderPath
    )
    
    If(Test-Path -Path $WordDocumentPath)
    {
        If(Test-Path -Path $ImageFolderPath)
        {
            $WordExtension = (Get-Item -Path $WordDocumentPath).Extension
            If($WordExtension -like ".doc" -or $WordExtension -like ".docx")
            {
                $ImageFiles = Get-ChildItem -Path $ImageFolderPath -Recurse -Include *.emf,*.wmf,*.jpg,*.jpeg,*.jfif,*.png,*.jpe,*.bmp,*.dib,*.rle,*.gif,*.emz,*.wmz,*.pcz,*.tif,*.tiff,*.eps,*.pct,*.pict,*.wpg
                
                If($ImageFiles)
                {
                    #Create the Word application object
                    $WordAPP = New-Object -ComObject Word.Application
                    $WordDoc = $WordAPP.Documents.Open("$WordDocumentPath")
                
                    Foreach($ImageFile in $ImageFiles)
                    {
                        $ImageFilePath = $ImageFile.FullName    
                        
                        $Properties = @{'ImageName' = $ImageFile.Name
                                        'Action(Insert)' = Try
                                                            {
                                                                $WordAPP.Selection.EndKey(6)|Out-Null
                                                                $WordApp.Selection.InlineShapes.AddPicture("$ImageFilePath")|Out-Null
                                                                
                                                                "Finished"
                                                            }
                                                            Catch
                                                            {
                                                                "Unfinished"
                                                            }
                                        }

                        $objWord = New-Object -TypeName PSObject -Property $Properties
                        $objWord
                    }

                    $WordDoc.Save()
                    $WordDoc.Close()
                    $WordAPP.Quit()#release the object
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordAPP)|Out-Null
                    Remove-Variable WordAPP
                }
                Else
                {
                    Write-Warning "There is no image in this '$ImageFolderPath' folder."
                }
            }
            Else
            {
                Write-Warning "There is no word document file in this '$WordDocumentPath' folder."
            }
        }
        Else
        {
            Write-Warning "Cannot find path '$ImageFolderPath' because it does not exist."
        }
    }
    Else
    {
        Write-Warning "Cannot find path '$WordDocumentPath' because it does not exist."
    }
    
    
}

# Constant Variables
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$templatePath = "$scriptDir\master_coverletter_template.docx"
$signaturePath = "$scriptDir\signature.png"
$outputPath = "$scriptDir\master_coverletter_output.docx"
$outputPathPDF = "$scriptDir\master_coverletter_output.pdf"

# Get info from temp.txt
#   name->value
$infoContents = Get-Content -Path "$scriptDir\..\scraper\temp.txt"

# Start Word Object
$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
$OpenFile = $Word.Documents.Open($templatePath)
$Content = $OpenFile.Content

# New variable for new text and variables to to replace the ones from the doc.
$newText = $Content.Text

# Adding employer info from temp.txt
Foreach ($line in $infoContents) {
    $lineSplit = $line -split "->"
    $newText = $newText  -replace "<$($lineSplit[0])>", $lineSplit[1]
}                         

# Make the modified text the new content and Save to new document
$Content.Text = $newText
$OpenFile.Saveas($outputPath)
$OpenFile.close()
$Word.Quit()

# Appending signature image
Add-OSCPicture -WordDocumentPath $outputPath -ImageFolderPath $signaturePath 

# Open file and save as PDF
$Word = New-Object -ComObject Word.Application
$Word.Visible = $False
$OpenFile = $Word.Documents.Open($outputPath)
$OpenFile.Saveas($outputPathPDF,  17)
$OpenFile.close()
$Word.Quit()

# Closing Message
Write-Output 'Generation Complete!'