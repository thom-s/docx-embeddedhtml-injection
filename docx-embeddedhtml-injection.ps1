# Parameters
$working_path = "C:\Users\example\Documents\" # Path of document
$original_docx = "test.docx"                 # Name of existing document (in path)
$html_block = ""                             # HTML you want to insert in the embeddedHTML document

# Open the document as an archive

Rename-Item -Path $working_path$original_docx -NewName "archive.zip"

Expand-Archive "$working_path\archive.zip" -DestinationPath "$working_path\archive"

Rename-Item -Path "$working_path\archive.zip" -NewName $original_docx

# Parse the XML file

[xml]$XmlDocument = Get-Content -Path "$working_path\archive\word\document.xml"

$XmlDocument.SelectNodes("//*").embeddedHtml

# Full node location in test document :  .p.r.drawing.inline.graphic.graphicData.pic.blipFill.blip.extLst.ext[1].webVideoPr.embeddedHtml
