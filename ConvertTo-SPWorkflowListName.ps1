
<#Hashtable that stores all list names and  list guids that is used for conversion
Example:

$listsMappings = @{
	"ListOneName" = "1356242A-D2FB-4865-A68A-0D91D38B4BC2";
	"ListTwoName" = "20e7aad9-d72e-4ee6-b66e-3d98598373ee"
}
#>
$listsMappings = @{}



function ConvertTo-SPWorkflowListName
{

  <#	
.SYNOPSIS
	Takes a folder that has .VWI files or a single .VWI file and converts the list guids in the workflow to list names in order to make the workflow portable.
	Also, has a "ConvertBack" parameter to convert a workflow  back to list guids. 
	The listMappings array needs to be configured inside the script file before running the script. 
	
.DESCRIPTION
	Takes a folder that has .VWI files or a single .VWI file and converts the list guids in the workflow to list names in order to make the workflow portable.
	Also, has a "ConvertBack" parameter to convert a workflow  back to list guids. 
	The listMappings array needs to be configured inside the script file before running the script.
	
.PARAMETER
	"-FileLocation" is the path to the .VWI file or folder that contains the .VWI files.
	
.PARAMETER
	"-ConvertBack" when present converts the list names back to a list guids. 
	
.EXAMPLE		
	Convert list guids to list names for a folder of .vwi files
	ConvertTo-SPWorkflowListName  -FileLocation C:/Workflows
	
	Convert list guids to list names for a single .vwi 
	ConvertTo-SPWorkflowListName  -FileLocation C:/Workflows/sample.vwi 	
	
.EXAMPLE
	Convert list names to list guids for a folder of .vwi files
	ConvertTo-SPWorkflowListName -FileLocation C:/Workflows -ConvertBack
	
	Convert list names to list guids for a single .vwi 
	ConvertTo-SPWorkflowListName  -FileLocation C:/Workflows/sample.vwi -ConvertBack	
	
.NOTES
	Function: ConvertTo-SPWorkflowListName
	Author: Alex Williamson

.LINK
	Github repo: https://github.com/alexcwilliamson/PowerShell
#>


  param(
    [Parameter(Mandatory = $true,
      HelpMessage = "Enter file or folder location of .vwi file
			- Use full path")]
    [string]$FileLocation,
    [Parameter(Mandatory = $false,
      HelpMessage = "Add to convert back to list guids
		")]
    [switch]$ConvertBack
  )
  if ($listsMappings.count -gt 0) {
    if (Test-Path $FileLocation -PathType container) {

      $files = Get-ChildItem $FileLocation

      if ($files.count -gt 0)
      {
        foreach ($file in $files)
        {
          if ($file -ne $Null) {
            if ($file.Extension -eq ".vwi")
            {
              try {
                Invoke-Convert $file $ConvertBack
              }
              catch {
                $Message = "[-] Failed to convert file: $($file.FullName) "
                Write-Error $Message
              }

            }
          }
        }
      }
    }
    else
    {

      $file = Get-ChildItem $FileLocation

      if ($file -ne $Null) {
        if ($file.Extension -eq ".vwi") {

          try {
            Invoke-Convert $file $ConvertBack
          }
          catch {
            $Message = "[-] Failed to convert file: $($file.FullName)"
            Write-Error $Message
          }
        }
      }

    }
  }
  else
  {
    $Message = "[-] Please configure the ListMappings variable inside the script file before running the script. "
    Write-Warning $Message
  }

}

function New-ZIP {

  <# 
 
.SYNOPSIS  
   Add files in a folder to a Zip folder using native feature in Windows 
    
.DESCRIPTION  
   Add files in a folder to a Zip folder using native feature in Windows 
 
.Parameter Source 
   Enter folder to be zipped - Use full path. 
   
.Parameter ZipFileName 
   Enter zip file name - if full path is not entered then parent of source is used. 

.EXAMPLE  
   New-ZIP -Source "c:\scripts\" -ZipFileName "x.zip" 
 
   Description 
   ----------- 
   Add folder "c:\scripts\" to "c:\x.zip" 
 
   (Notice that ZIP file is placed in source folder's parent folder)
   
.EXAMPLE  
   New-ZIP c:\scripts x.zip 
 
   Description 
   ----------- 
   Add folder c:\scripts to c:\x.zip 
 
   (Notice this is a simplified command line) 
   - The ZIP file is placed in source folder's parent folder 
   - There is no trailing slash on source folder  
   - No parameter names specified 
   - No quotes because there are no spaces in path
   
.NOTES  
   NAME:      New-ZIP 
   AUTHOR:    Matthew Painter 
   CREDITS:   Modified version of Matthew Painters script at 
   
.LINK
   http://gallery.technet.microsoft.com/ScriptCenter/en-us/e093e2d0-672b-402b-a33f-568610cc4bb7 
    

 
#>

  [CmdletBinding()]

  param(

    [Parameter(
      ValueFromPipeline = $False,
      Mandatory = $true,
      HelpMessage = "Enter folder to be zipped 
			- Use full path")]
    [string]$Source,

    [Parameter(
      ValueFromPipeline = $False,
      Mandatory = $true,
      HelpMessage = "Enter zip file name 
			- if full path is not entered then parent of source is used")]
    [ValidatePattern('.\.zip$')]
    [string]$ZipFileName

  )

  $zipHeader = [char]80 + [char]75 + [char]5 + [char]6 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0 + [char]0

  if ((Test-Path $ZipFileName) -eq $FALSE)
  {
    Add-Content $ZipFileName -Value $zipHeader
  }
  else
  {
    Remove-Item $ZipFileName -Recurse
    Add-Content $ZipFileName -Value $zipHeader
  }

  $ExplorerShell = New-Object -ComObject 'Shell.Application'
  $zipFileItems = Get-ChildItem $Source

  foreach ($zipFile in $zipFileItems)
  {
    $ExplorerShell.Namespace($ZipFileName.ToString()).CopyHere($zipFile.FullName)
    Start-Sleep -Milliseconds 500
  }

}

function Expand-ZIPFile ($file,$destination)
{
  #Unzips the files in the $file directory and copy's them to the $destination directory
  $shell = New-Object -com shell.application
  $zip = $shell.Namespace($file)
  New-Item ($destination) -Type directory

  foreach ($item in $zip.items())
  {
    $shell.Namespace($destination).CopyHere($item)
  }
}




function Invoke-Convert ($file,$ConvertBack)
{
  #Rename's vwi file to a zip file
  Rename-Item $file.FullName ($file.Name + ".zip")
  #Calls a function that unzips the renamed zip file to a new folder 
  Expand-ZIPFile ($file.FullName + ".zip") ($file.FullName)
  #Gets all the files in the unzipped folder and loops through each one
  $zipFiles = Get-ChildItem $file.FullName

  foreach ($zipFile in $zipFiles)
  {
    if ($zipFile.Name -eq "workflow.xoml")
    {
      #Gets the text in the xoml file
      $xomlContent = Get-Content $zipFile.FullName
      #Calls the function to convert the list guid to list name if ConvertBack is false
      if ($ConvertBack -eq $false)
      {
        $xomlContent = Set-ListNameContent $xomlContent
      }
      #Calls the function to convert the list name to list guid if ConvertBack is true
      else
      {
        $xomlContent = Set-ListGuidContent $xomlContent
      }
      #Writes file contents to unzipped file
      $xomlContent > $zipFile.FullName

    }
    #Deletes config file to enable the workflow to be imported in SharePoint Designer
    elseif ($zipFile.Name -eq "workflow.xoml.wfconfig.xml")
    {
      Remove-Item $zipFile.FullName
    }
  }
  #Deletes zip file
  Remove-Item ($file.FullName + ".zip") -Recurse
  #Converts folder with the modified xoml file to a zip file 
  New-ZIP -Source ($file.FullName) -ZipFileName ($file.FullName + ".zip")
  #Removes folder with the modified xoml file
  Remove-Item ($file.FullName) -Recurse
  #Renames the zip file to a .vwi file
  Rename-Item ($file.FullName + ".zip") $file.Name

}


function Set-ListNameContent ($xomlContent)
{
  #Function that  uses the $listMappings array to replace the  list guids to list names in the xoml file.

  foreach ($item in $listsMappings.GetEnumerator())
  {
    $xomlContent = $xomlContent.replace("ListId=`"{}{$($item.Value)}`"","ListId=`"$($item.Key)`"")
    $xomlContent = $xomlContent.replace("ExternalListId=`"{}{$($item.Value)}`"","ExternalListId=`"$($item.Key)`"")
  }

  return $xomlContent
}

function Set-ListGuidContent ($xomlContent)
{
  #Function that  uses the $listMappings array to replace the  list names to list guids in the xoml file.

  foreach ($item in $listsMappings.GetEnumerator())
  {
    $xomlContent = $xomlContent.replace("ListId=`"$($item.Key)`"","ListId=`"{}{$($item.Value)}`"")
    $xomlContent = $xomlContent.replace("ExternalListId=`"$($item.Key)`"","ExternalListId=`"{}{$($item.Value)}`"")

  }

  return $xomlContent
}




