# Parameters
$working_path = "C:\Users\example\Documents\" # Path of document
$original_docx = "test.docx"                  # Name of existing document (in path)
$html_block = ""                              # HTML you want to insert in the embeddedHTML document

# Open the document as an archive

Rename-Item -Path $working_path$original_docx -NewName "archive.zip"

Expand-Archive "$working_path\archive.zip" -DestinationPath "$working_path\archive"

Rename-Item -Path "$working_path\archive.zip" -NewName $original_docx

# Parse the XML file

$xml = New-Object XML
$xml.Load("$working_path\archive\word\document.xml")
$xml.SelectNodes("//*") | ForEach-Object {
    if($_.embeddedHtml){
        $_.SetAttribute("embeddedHtml",$html_block)
    }
}
$xml.Save("$working_path\archive\word\document.xml")

# Close back the file

Compress-Archive -Path "$working_path\archive\*" -DestinationPath "$working_path\archive.zip"

Remove-Item -Path "$working_path\archive" -Force -Recurse

Rename-Item -Path "$working_path\archive.zip" -NewName "output.docx"
