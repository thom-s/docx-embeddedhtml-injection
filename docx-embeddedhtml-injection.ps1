function Inject-Docx{
    param(
        [string]$Path,                           # Full path of document
        [string]$HtmlBlock,                      # HTML you want to insert in the embeddedHTML document
        [string]$DestinationName = "output.docx" # Optional : File name of the output file
    )

    # Split $Path argument into variables
    $working_path = Split-Path $Path -Parent      # Path of document
    $original_docx = Split-Path $Path -Leaf       # Name of existing document (in path)

    # Open the document as an archive
    Rename-Item -Path $Path -NewName "archive.zip"                                      # Rename document to zip file 
    Expand-Archive "$working_path\archive.zip" -DestinationPath "$working_path\archive" # Extract the zip file into a new folder
    Rename-Item -Path "$working_path\archive.zip" -NewName $original_docx               # Rename the zip file to its original name (docx file)

    # Parse the XML file
    $xml = New-Object XML
    $xml.Load("$working_path\archive\word\document.xml") # Open the document.xml file (where the embeddedHtml tag is located)
    $xml.SelectNodes("//*") | ForEach-Object {           # Loop through all XML nodes
        if($_.embeddedHtml){                             # If the node has an embeddedHtml attribute
            $_.SetAttribute("embeddedHtml",$HtmlBlock)   # Replace the attribute value with the $HtmlBlock parameter
        }
    }
    $xml.Save("$working_path\archive\word\document.xml") # Save the XML

    # Compress the extracted folder into a new file
    Compress-Archive -Path "$working_path\archive\*" -DestinationPath "$working_path\archive.zip" # Compress the folder into a new zip file
    Remove-Item -Path "$working_path\archive" -Force -Recurse                                     # Delete the folder
    Rename-Item -Path "$working_path\archive.zip" -NewName $DestinationName                       # Rename the new zip file
}
