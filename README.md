# PowerResizer

PowerShell script which can rotate, compress and scale down input file images. Its main purpose is to prepare a bunch of Images in a given folder for web-publishing.

**ATTENTION**: always keep in mind that PowerResizer will **always overwrite** your input files! So you probably do **not** want to run this script on t one-and-only-copy of your photo-library!

Uses external Programs for conversions:
PhotoMolo:
http://www.ktverkko.fi/~msmakela/software/photomolo/photomolo.html
ImageMagick:
https://www.imagemagick.org/script/index.php


##DESCRIPTION

Script will optimize common image formats, result will always be a progressive JPG.
Depending on the input file, 2 different optimization parameters will be set
for images larger than a given resolution and images with a smaller resolution.
Define your requirements in script section "User Settings".
Script can work recursively through a given directory, defaults to non-recursive.

ATTENTION: this script will overwrite the original files, and also remove non-JPG
input files!! Handle with care!

##PARAMETER InputPath
Input Path might be a single file or a folder.
If InputPath is a Folder, all files with supported extensions will be treated.

##PARAMETER Recurse
Only supported when InputPath is a folder. Work recursively through all subfolders.
Handle with care!

##EXAMPLE

    C:\PS> .\PowerResizer.ps1 -InputPath 'c:\Pics\to convert\convertme.bmp'

##EXAMPLE

    C:\PS> .\PowerResizer.ps1 -InputPath 'c:\Pics\to convert\recursively' -Recurse

##HINT

You might want to place a CMD-wrapper script in your "Send To" folder which is located in:  
%APPDATA%\Microsoft\Windows\SendTo  
for recent Windows OSes.  
The Wrapper script should contain:  
    @powershell.exe -ExecutionPolicy unrestricted -file "C:\PathTo\PowerResizer.ps1" %1

It will allow you to Use the Windows Explorers "Send To" context Menu to send an image file or folder through the PowerResizer.

Have fun,
Juergen
