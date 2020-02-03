# Global Variables
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$template = "$scriptDir\master_coverletter_template.docx"
$output = "$scriptDir\master_coverletter_output.docx"

# Start Word Object
$Word = New-Object -ComObject Word.Application
$Word.Visible = $False

$OpenFile = $Word.Documents.Open($template)
$Content = $OpenFile.Content

# New variable for new text and variables to to replace the ones from the doc.
$newText = $Content.Text

$newText = $newText  -replace '<name>', 'billy'

# Make the modified text the new content and Save to new document
$Content.Text = $newText
$OpenFile.Saveas($output)
$OpenFile.close()