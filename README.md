# .SYNOPSIS
    Extracts EXIF metadata from images to Word
# .DESCRIPTION
    Extracts EXIF metadata from all image files located in $PhotosFolderPath, and saves the images and corresponding EXIF metadata to a Word Document.
# .EXAMPLE
    PS C:\> .\Extract-ImageExifDataToWord.ps1 -PhotosFolderPath ".\example"
    Extracts EXIF data from all images in ".\example", and saves the images and corresponding EXIF metadata to a Word Document named "EXIF_Photo_Data.docx" in the same folder.
    
    [OPTIONAL]
    PS C:\> .\Extract-ImageExifDataToWord.ps1 -PhotosFolderPath ".\example" -WordDocumentName "EXIF-ImageData-Project"
    Extracts EXIF data from all images in ".\example", and saves the images and corresponding EXIF metadata to a Word Document named "EXIF-ImageData-Project.docx" in the same folder.

# .INPUTS
    -PhotosFolderPath -> Folder path location for the image files you'd like to extract EXIF metadata from.
    -WordDocumentName -> [OPTIONAL] Name for the Word document where the EXIF metadata and images will be saved in. Default name is: "EXIF_Photo_Data"
# .OUTPUTS
    Word Document with all the images from $PhotosFolderPath, combined with the corresponding EXIF metadata information located in a table beneath each image. 
# .NOTES
    Thanks to https://community.spiceworks.com/people/mohand    -> https://community.spiceworks.com/topic/1250688-powershell-script-to-read-metadata-info-from-pictures
    Thanks to https://github.com/EvotecIT                       -> https://github.com/EvotecIT/PSWriteWord

