<#
.SYNOPSIS
    Extracts EXIF metadata from images to Word
.DESCRIPTION
    Extracts EXIF metadata from all image files located in $PhotosFolderPath, and saves the images and corresponding EXIF metadata to a Word Document.
.EXAMPLE
    PS C:\> .\Extract-ImageExifDataToWord.ps1 -PhotosFolderPath ".\example"
    Extracts EXIF data from all images in ".\example", and saves the images and corresponding EXIF metadata to a Word Document named "EXIF_Photo_Data.docx" in the same folder.
    
    [OPTIONAL]
    PS C:\> .\Extract-ImageExifDataToWord.ps1 -PhotosFolderPath ".\example" -WordDocumentName "EXIF-ImageData-Project"
    Extracts EXIF data from all images in ".\example", and saves the images and corresponding EXIF metadata to a Word Document named "EXIF-ImageData-Project.docx" in the same folder.

.INPUTS
    -PhotosFolderPath -> Folder path location for the image files you'd like to extract EXIF metadata from.
    -WordDocumentName -> [OPTIONAL] Name for the Word document where the EXIF metadata and images will be saved in. Default name is: "EXIF_Photo_Data"
.OUTPUTS
    Word Document with all the images from $PhotosFolderPath, combined with the corresponding EXIF metadata information located in a table beneath each image. 
.NOTES
    Thanks to https://community.spiceworks.com/people/mohand    -> https://community.spiceworks.com/topic/1250688-powershell-script-to-read-metadata-info-from-pictures
    Thanks to https://github.com/EvotecIT                       -> https://github.com/EvotecIT/PSWriteWord
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, HelpMessage = "Enter the folder path where the photos are located. Example: 'C:\Photos\'")]
    [string] $PhotosFolderPath,

    [Parameter(Mandatory=$False, HelpMessage = "Enter the Word Document name. Default: 'EXIF_Photo_Data'")]
    [string] $WordDocumentName = "EXIF_Photo_Data"
)

$ErrorActionPreference = "Stop"

try{
#region Preparation

    # Variables
    # These variables control the Word Document general formatting. For more extensive format control, please review the documentation for PSWriteWord (https://github.com/EvotecIT/PSWriteWord) 
    [int]$WordImageWidth = 500 # Value is in Pixels. Heigth of the image will be automatically calculated based upon Width and original proportions.
    [double]$WordImageTitleFontSize = 12
    [double]$WordImageTableFontSize = 8

    # Install module if not found and load it
    if (Test-Path "$($PSScriptRoot)\modules\PSWriteWord\1.1.11\PSWriteWord.psm1") { import-Module "$($PSScriptRoot)\modules\PSWriteWord\1.1.11\PSWriteWord.psm1" }
    else { Install-Module -Name PSWriteWord -Scope CurrentUser -Force ; Import-Module PSWriteWord }

    # Load Functions
    $functions = Get-ChildItem -Path "$($PSScriptRoot)\functions"
    foreach ($function in $functions) {. "$($PSScriptRoot)\functions\$($function)" }

    # Sanitize Parameters
    if ( !($PhotosFolderPath.EndsWith("\")) )   { $PhotosFolderPath = $PhotosFolderPath + "\" }                 # Add trailing slash if not present
    if ( $WordDocumentName.EndsWith(".docx") )  { $WordDocumentName = $WordDocumentName.Replace(".docx", "") }  # Remove .docx from parameter
    if ( $WordDocumentName.EndsWith(".doc") )   { $WordDocumentName = $WordDocumentName.Replace(".doc", "") }   # Remove .doc from parameter

    # Ask user for new folder location if it does not exist
    while ( !(Test-Path $PhotosFolderPath) ) { $PhotosFolderPath = Read-Host -Prompt "The photos folder path does not exist, please enter a valid folder path where the photos are located. Example: 'C:\Fotos\'" }

#endregion
}
catch{
    Write-Host "There was an error during preparation."
    Read-Host -Prompt "Press any key to exit.."
    throw
}

try{
#region Execution

    # Create new Word document
    $WordDocument = $null
    $WordDocument = New-WordDocument "$($PhotosFolderPath)$($WordDocumentName).docx"

    # Get Photo EXIF data
    $ExifData = Get-FileMetaData -Folder $PhotosFolderPath
    $image = $null

    foreach ($image in $ExifData)
    {
        # Calculate photo proportions
        $WordImageProportion = ($image.Breedte).remove(0,1).Replace(" pixels","") / ($image.Hoogte).Remove(0,1).Replace(" pixels","")
        $WordImageWidth = 500
        $WordImageHeigth = $WordImageWidth / $WordImageProportion

        # Add content to Word Document
        Add-WordText -WordDocument $WordDocument -Text $image.Naam -FontSize $WordImageTitleFontSize -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $image.pad -ImageWidth $WordImageWidth -ImageHeight $WordImageHeigth -Alignment left -Supress $true
        Add-WordParagraph -WordDocument $WordDocument -Supress $true

        # Reorganize EXIF data
        $array = @()
        $props = $image | Get-Member -MemberType NoteProperty 

        foreach ($prop in $props)
        {
            if ($image.($prop.name).Length -gt 0)
            { 
                $array += [pscustomobject]@{ Name = $prop.name; Value = $image.($prop.name) } 
            }
        }

        # Add EXIF data to Word
        $table = Add-WordTable -WordDocument $WordDocument -DataTable $array  -FontSize $WordImageTableFontSize -Design LightShading -ContinueFormatting -AutoFit Contents -BreakPageAfterTable -Supress $true
    }

    # Save all changes to Word document
    Save-WordDocument -WordDocument $WordDocument -Language 'nl-nl' -KillWord -OpenDocument

    Write-Host "All done! Word Document saved at: $($PhotosFolderPath)$($WordDocumentName).docx"

#endregion
}
catch{
    Write-Host "There was an error while creating the Word document with EXIF data."
    Write-Host $Error[0]
    Read-Host -Prompt "Press any key to exit.."
    throw
}

Read-Host -Prompt "Press any key to exit.."
