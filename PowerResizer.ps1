################################ PowerResizer #################################
############ Resize, convert, rotate images and store as JPG ##################
###############################################################################
# Requires external programs:
# PhotoMolo for Auto-Rotating images:
# http://www.ktverkko.fi/~msmakela/software/photomolo/photomolo.html
# ImageMagick for resizing:
# https://www.imagemagick.org/script/index.php

# Only tested in Windows 10
#Requires -Version 5

<#
.SYNOPSIS

Resizes, converts and compresses input images and saves them as JPG files.

.DESCRIPTION

Script will optimize common image formats, result will always be a progressive JPG.
Depending on the input file, 2 different optimization parameters will be set
for images larger than a given resolution and images with a smaller resolution.
Define your requirements in script section "User Settings".
Script can work recursively through a given directory, defaults to non-recursive.

ATTENTION: this script will overwrite the original files, and also remove non-JPG
input files!! Handle with care!

.PARAMETER InputPath
Input Path might be a single file or a folder.
If InputPath is a Folder, all files with supported extensions will be treated.

.PARAMETER Recurse
Only supported when InputPath is a folder. Work recursively through all subfolders.
Handle with care!

.PARAMETER AddPrefix
Will add the directory name, where the handled photo is located, as file name prefix.
Files which already start with the new prefix will not be renamed.
Result will be DIRNAME + "_" + Original-Filename
Main purpose is to get some overview in flat WordPress Media library.

.PARAMETER Quality
Alternatively to compress photos to a target file size, you can use this param
to manually set a desired JPEG compression level. Values 0 to 100 are accepted, value 0
is the default which means target file size as defined in "User Settings".

.EXAMPLE

C:\PS> .\PowerResizer.ps1 -InputPath 'c:\Pics\to convert\convertme.bmp'

.EXAMPLE

C:\PS> .\PowerResizer.ps1 -InputPath 'c:\Pics\to convert\recursively' -Recurse

.EXAMPLE

C:\PS> .\PowerResizer.ps1 -InputPath 'c:\Pics\EventXY' -AddPrefix
Result: Photo "abc.jpg" will be renamed to "EventXY_abc.jpg"

#>

#region ###################### Script Parameters ##############################
[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$true,Position=1)]
    [ValidateScript({if (Test-Path $_){return $true} else {throw "InputPath does not exist."}})]
    [string]$InputPath,
    [Parameter(Mandatory=$false)]
    [switch]$Recurse,
    [Parameter(Mandatory=$false)]
    [switch]$AddPrefix,
    [Parameter(Mandatory=$false)]
    [ValidateRange(0,100)]
    [Int]$Quality = 0
)
#endregion ####################################################################


# Add required assemblies
Add-Type -Assembly System.Drawing


#region ##################### User Settings ###################################

# Input Image file name extensions that will be handled
$SupportedInputFileExt = @('*.jpg','*.jpeg','*.png','*.tif','*.tiff','*.bmp')

# Minimum file size for Images to be edited by ExifIron
# ExifIron can only rotate and optimize camera / handy photos, which are usually quite large..
[Int]$MinExifIronImgSizeMB = 3

# Default desired JPG output file sizes depending on original dimesnions of the input file
# Input Pictures with small dimensions will not be resized but only compressed (if necessary)
# to the desired target file size (in kB)
[Int]$SmallPicInMaxLongSidePx = 1000
[Int]$SmallPicOutTargetSizeKB = 100
# Large input pictures will be resized to desired dimension (the longer side, aspect ratio remains)
# and stored with a compression to reach the desired output file size (in kB)
[Int]$LargePicOutTargetSizeKB = 250
[Int]$LargePicOutScaledLongSidePx = 1280

#endregion ####################################################################

#region ##################### System Settings #################################

# Path to external binaries
$ImgMgk="C:\Program Files\ImageMagick-7.0.8-Q8\magick.exe"
$ExifIron="C:\Program Files (x86)\photomolo-win32-1_2_5\exifiron.exe"

#endregion ####################################################################

#region #########################  Functions  #################################

# Function for ImageMagick resizing and format conversion
function mogrify
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({if (Test-Path $_ -PathType Leaf){return $true} else {throw "File parameter missing or not a file."}})]
        [string]$file,
        [Parameter(Mandatory=$false)]
        [Int]$LongSidePx=0,
        [Parameter(Mandatory=$false)]
        [Int]$Quality=0,
        [Parameter(Mandatory=$false)]
        [Int]$SizeKB = "500"
    )

    # Helpers
    $RemoveInputFile = $false

    # Get file properties
    $InputFile = Get-Item $file

    $FullIMParams = @()
    # Inline modification of the picture
    $FullIMParams += "mogrify"
    # Create Progressive JPGs
    $FullIMParams += "-interlace Plane"
    # Add resize opions if LongSidePx parameter is given
    if ( $LongSidePx -gt 0 )
    {
        $FullIMParams += $("-resize $($LongSidePx)x$($LongSidePx)" + '>')
    }
    # If Input file is not a JPG, convert to JPG
    if ( -not ($InputFile.Extension -imatch "jpg|jpeg") )
    {
        Write-Host "Converting $($InputFile.Extension) Inputfile to JPG.." -ForegroundColor Green
        $FullIMParams += "-format jpg"
        $RemoveInputFile = $true
    }
    # Set desired compression quality or target file size
    if ($Local:Quality -eq 0)
    {
        # No quality param given, use user defined target file size
        $FullIMParams += "-define jpeg:extent=$($SizeKB)KB"
    }
    else
    {
        # use given quality value
        $FullIMParams += "-quality $Local:Quality"
    }

    # Add file with full path
    $FullIMParams += "`"$($InputFile.FullName)`""

    #Write-Host "DEBUG IM Params:" -ForegroundColor Red
    #$FullIMParams | Out-String
    #pause

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $ImgMgk
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = $FullIMParams
    $pinfo.CreateNoWindow = $true
    $IMproc = New-Object System.Diagnostics.Process
    $IMproc.StartInfo = $pinfo
    # Start ImageMagick and wait until finished
    $IMproc.Start() | Out-Null
    $IMproc.WaitForExit()

    if ( $IMproc.ExitCode -ne 0 )
    {
        # Error?
        $IMstdout="$($IMproc.StandardOutput.ReadToEnd())"
        $IMstderr="$($IMproc.StandardError.ReadToEnd())"
        Write-Host "ImageMagick STDOUT:`n$IMstdout" -ForegroundColor Yellow
        Write-Host "ImageMagick STDERR:`n$IMstderr" -ForegroundColor Red
        Write-Error "ImageMagick failed handling $($InputFile.Name)!" -ErrorAction Stop
    }
    else
    {
        if ($RemoveInputFile)
        {
            Write-Host "Removing original input file $($InputFile.Name).." -ForegroundColor Yellow
            Remove-Item -Force $InputFile
        }
    }
}

# Function to rotate and cleanup Camera Photos with ExifIron
function rotate
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({if (Test-Path $_ -PathType Leaf){return $true} else {throw "File parameter missing or not a file."}})]
        [string]$file,
        [Parameter(Mandatory=$false)]
        [Int]$MinSizeMB = "2"
    )
    

    # Get file properties
    $InputFile = Get-Item $file

    # If File is smaller than $MinSizeMB, it is probably not a camera photo, so don't touch it
    if (($InputFile.Length/1MB) -lt $MinSizeMB)
    {
        Write-Host "Skipping file (probably not a suitable camera photo)." -ForegroundColor Yellow
        return
    }

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $ExifIron
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = $InputFile.FullName
    $pinfo.CreateNoWindow = $true
    $EIproc = New-Object System.Diagnostics.Process
    $EIproc.StartInfo = $pinfo
    # Start ExifIron and wait until finished
    $EIproc.Start() | Out-Null
    $EIproc.WaitForExit()

    if ( $EIproc.ExitCode -ne 0 )
    {
        # Error?
        $EIstdout="$($EIproc.StandardOutput.ReadToEnd())"
        $EIstderr="$($EIproc.StandardError.ReadToEnd())"
        Write-Host "ExifIron STDOUT:`n$EIstdout" -ForegroundColor Yellow
        Write-Host "ExifIron STDERR:`n$EIstderr" -ForegroundColor Red
        Write-Error "ExifIron failed handling $($InputFile.Name)!" -ErrorAction Stop
    }
}

# Function to get long side of an Image in pixel
function get-LongsideImagePx
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateScript({if (Test-Path $_ -PathType Leaf){return $true} else {throw "File parameter missing or not a file."}})]
        [string]$file
    )

    try { $ImageProps = [System.Drawing.Image]::FromFile($file) }
    catch { Write-Error "Failed to get image properties from $($file.FullName)! Not a valid image file?" -ErrorAction Stop }

    # Get required properties and dispose $ImageProps to release the file handle (else we can't modify it!)
    [Int]$ImgWidth = $ImageProps.Width
    [Int]$ImgHeight = $ImageProps.Height
    $ImageProps.Dispose()
    
    if ( $ImgWidth -ge $ImgHeight )
    {
        return $ImgWidth
    }
    else
    {
        return $ImgHeight
    }
}

#endregion


#region ############################## Main Script #####################################

# Test if external binaries are available
if (-not (Test-Path $ImgMgk -PathType Leaf))
{
    Write-Error "Required ImageMagick binary not found! Check System Settings!" -ErrorAction Stop
}
if (-not (Test-Path $ExifIron -PathType Leaf))
{
    Write-Error "Required ExifIron (PhotoMolo) binary not found! Check System Settings!" -ErrorAction Stop
}


# Analyze InputPath and extract files to handle
if (Test-Path $InputPath -PathType Container)
{
    # InputPath is a directory, add * for get-childitem
    $InputPath = $($InputPath + '\*')
    # Handle every supported file in fodler (recursively, if requested)
    if ($Recurse)
    {
        $HandleItems = Get-ChildItem -Path $InputPath -Include $SupportedInputFileExt -Recurse
    }
    else
    {
        $HandleItems = Get-ChildItem -Path $InputPath -Include $SupportedInputFileExt
    }
    if ($HandleItems.Count -eq 0)
    {
        Write-Error "No Images to handle in path $InputPath!" -ErrorAction Stop
    }
}
else
{
    # Input is a file, check if supported
    $HandleItems = Get-Item $InputPath
    if (-not $SupportedInputFileExt -imatch $HandleItems.Extension)
    {
        Write-Error "Input file $($HandleItems.FullName) is not supported!" -ErrorAction Stop
    }
}

# Start the magic
$HandleItems | ForEach-Object {

    Write-Host "PowerResizer start working on $($_.FullName).." -ForegroundColor Green
    # Get file properties
    $InputFile = $_
    # Get Pixelcount of the longer image side
    [Int]$PixCount = get-LongsideImagePx -file $InputFile.FullName

    Write-Host "File Size: $([math]::Round($InputFile.Length / 1KB)) kB`nLong Side: $PixCount px" -ForegroundColor Cyan

    # Rename input file if AddPrefix $true
    if ($AddPrefix)
    {
        $LeafDirName = $(Split-Path -Leaf $InputFile.Directory)
        $NewFileName = $($LeafDirName + '_' + $InputFile.Name)

        # Rename file only if it doesn't already have the right prefix
        if ($InputFile.Name -notmatch "^$LeafDirName")
        {
            Write-Host "Renaming file to: $NewFileName" -ForegroundColor Yellow
            Rename-Item -Path $InputFile.FullName -NewName $NewFileName -ErrorAction Stop
            # Update InputFile
            $InputFile = Get-Item ($InputFile.Directory.ToString() + '\' + $NewFileName)
        }
    }


    # ExifIron only for JPGs
    if ($InputFile.Extension -imatch ".jpg")
    {
        Write-Host "Optimizing and rotating image with ExifIron.." -ForegroundColor Green
        rotate -file $InputFile.FullName -MinSizeMB $MinExifIronImgSizeMB
    }

    # ImageMagick
    # Check if we have a small or a large pic
    if ($PixCount -gt $SmallPicInMaxLongSidePx)
    {
        if ($Quality -ne 0)
        {
            Write-Host "Resizing and compressing large image file with ImageMagick (Target Quality $Quality)" -ForegroundColor Green
            mogrify -file $InputFile.FullName -LongSidePx $LargePicOutScaledLongSidePx -Quality $Quality
        }
        else
        {
            Write-Host "Resizing and compressing large image file with ImageMagick (Targetsize $($LargePicOutTargetSizeKB)kB)" -ForegroundColor Green
            mogrify -file $InputFile.FullName -LongSidePx $LargePicOutScaledLongSidePx -SizeKB $LargePicOutTargetSizeKB
        }
    }
    else
    {
        if ($Quality -ne 0)
        {
            Write-Host "Compressing small image file with ImageMagick (Target Quality $Quality)" -ForegroundColor Green
            mogrify -file $InputFile.FullName -Quality $Quality
        }
        else
        {
            Write-Host "Compressing small image file with ImageMagick (Targetsize $($SmallPicOutTargetSizeKB)kB)" -ForegroundColor Green
            mogrify -file $InputFile.FullName -SizeKB $SmallPicOutTargetSizeKB
        }
    }
    Write-Host "PowerResizer successfully handled $($InputFile.Name)." -ForegroundColor Green
    Write-Host "----------------------------------------------------------------------------"
}