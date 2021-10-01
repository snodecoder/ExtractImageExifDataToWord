<#
.SYNOPSIS
    Retrieves metadata from image files
.DESCRIPTION
    Retrieves metadata from image files located in the provided folder path
.EXAMPLE
    PS C:\> . .\functions\Get-FileMetaData.ps1 # Imports function
    PS C:\> $MetaDataObject = Get-FileMetaData -Folder ".\example"
    Retrieves metadata for all image files located in folder .\example
.INPUTS
    -Folder -> Folder path for the location of the image files.
.OUTPUTS
    Array of objects, with an object per file, filled with the image metadata for that file.
.NOTES
    
#>
Function Get-FileMetaData
{
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage="Enter the folder path where the photos are located.")]
        [string[]] $Folder
    )

    $array = @()
    foreach ($sFolder in $folder)
    {
        $a = 0
        $objShell = New-Object -ComObject Shell.Application
        $objFolder = $objShell.namespace($sFolder)

        foreach ($File in $objFolder.items())
        {
            # Only continue if File has image extension
            if ($File.Name -like "*.png" -or $File.Name -like "*.jpg" -or $File.Name -like "*.jpeg" -or $File.Name -like "*.gif" -or $File.Name -like "*.tif" -or $File.Name -like "*.tiff" -or $File.Name -like "*.bmp") 
            {
                $MetaDataObject = New-Object System.Object
                
                for ($a ; $a -le 266; $a++)
                {
                    if ($objFolder.getDetailsOf($File, $a))
                    {
                        $property = $objFolder.getDetailsOf($objFolder.items, $a)
                        $value = $objFolder.getDetailsOf($File, $a)

                        If (($Value -ne $null) -and ($Value -ne '')) 
                        {
                            $MetaDataObject | Add-Member -MemberType NoteProperty -Name $Property -Value $Value
                        }
                    } #end if
                } #end for

                $a = 0           
                $array += $MetaDataObject
            }
        } #end foreach $file       
    } #end foreach $sfolder

    return $array
} #end Get-FileMetaData
