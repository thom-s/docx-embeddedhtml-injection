function Inject-Docx{
    param(
        [string]$Path,      # Full path of document
        [string]$HtmlBlock, # HTML you want to insert in the embeddedHTML document
        [string]$DestinationName = "output.docx"
    )

    $working_path = Split-Path $Path -Parent     # Path of document
    $original_docx = Split-Path $Path -Leaf       # Name of existing document (in path)

    # Open the document as an archive

    Rename-Item -Path $Path -NewName "archive.zip"

    Expand-Archive "$working_path\archive.zip" -DestinationPath "$working_path\archive"

    Rename-Item -Path "$working_path\archive.zip" -NewName $original_docx

    # Parse the XML file

    $xml = New-Object XML
    $xml.Load("$working_path\archive\word\document.xml")
    $xml.SelectNodes("//*") | ForEach-Object {
        if($_.embeddedHtml){
            $_.SetAttribute("embeddedHtml",$HtmlBlock)
        }
    }
    $xml.Save("$working_path\archive\word\document.xml")

    # Close back the file

    Compress-Archive -Path "$working_path\archive\*" -DestinationPath "$working_path\archive.zip"

    Remove-Item -Path "$working_path\archive" -Force -Recurse

    Rename-Item -Path "$working_path\archive.zip" -NewName $DestinationName
}
