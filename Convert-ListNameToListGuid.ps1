
<#Hashtable that stores all list name's and  list guid's that is used in the ConvertTo-ListName and ConvertTo-ListGuid functions 
Example:

$listsMappings = @{
"ListOneName" = "1356242A-D2FB-4865-A68A-0D91D38B4BC2";
"ListTwoName" = "20e7aad9-d72e-4ee6-b66e-3d98598373ee";

}
#>
$listsMappings = @{}


function ConvertTo-ListName
{

  <#
.SYNOPSIS

Takes either one file or a folder and converts the Sharepoint webparts located in the file(s) from list guids to list names in order to make the webpart portable. The ListMappings array needs to be configured inside the script file. 

Function: ConvertTo-ListName
Author: Alex Williamson

.DESCRIPTION

Takes either one file or a folder and converts the Sharepoint webparts ilocated in the file(s) from list guids to list names in order to make the webpart portable. The ListMappings array needs to be configured inside the script file.

.EXAMPLE

ConvertTo-ListName -FileLocation "C:/Sitepages" or ConvertTo-ListName -FileLocation "C:/Sitepages/Home.aspx"

.NOTES

.LINK


Github repo: https://github.com/clymb3r/PowerShell
#>


  param(
    [Parameter(Mandatory = $true)]
    $FileLocation
  )

  if (Test-Path $FileLocation -PathType container) {

    $files = Get-ChildItem $FileLocation
    if ($files.count -gt 0)
    {
      foreach ($file in $files)
      {
        if ($file -ne $null)
        {
          try {
            $htmlContent = Get-Content $file.FullName
            $htmlContent = Set-ListNameContent ($htmlContent);
            $htmlContent > $file.FullName
          }
          catch {

            $message = "[-] Failed to convert file to ListName: $($file.FullName) "
            Write-Error $message

          }
        }
      }
    }

  }
  else
  {

    $htmlContent = Get-Content $FileLocation;
    if ($htmlContent -ne $null)
    {
      try {
        $htmlContent = Set-ListNameContentt ($htmlContent);
        $htmlContent > $FileLocation
      }
      catch {
        $message = "[-] Failed to convert file to ListName: $file "
        Write-Error $message

      }
    }

  }

}


function Set-ListNameContent ($htmlContent)
{
  #Function that  uses the $ListMappings array to replace the contents of the file from a list guid  to list name.

  if ($listsMappings.count -gt 0) {
    foreach ($Item in $listsMappings.GetEnumerator())
    {
      $htmlContent = $htmlContent.replace("ListId=`"{$($Item.Value)}`"","")
      $htmlContent = $htmlContent.replace("ListName=`"{$($Item.Value)}`"","ListDisplayName=`"$($Item.Key)`"")
      $htmlContent = $htmlContent.replace("DefaultValue=`"$($Item.Value)`"","DefaultValue=`"$($Item.Key)`"")
    }
  }
  else
  {
    $message = "[-] Please configure the ListMappings variable before running the script "
    Write-Warning  $message
    exit
  }

  $htmlContent = $htmlContent.replace("ListID","ListName")
  return $htmlContent;
}


function ConvertTo-ListGuid
{


  <#
.SYNOPSIS

Takes either one file or a folder and converts the Sharepoint webparts located in the file(s) from list guids to list names. The ListMappings array needs to be configured inside the script file. 

Function: ConvertTo-ListName
Author: Alex Williamson


.DESCRIPTION

Takes either one file or a folder and converts the Sharepoint webparts located in the file(s) from list guids to list names. The ListMappings array needs to be configured inside the script file. 

.EXAMPLE

ConvertTo-ListGuid -FileLocation "C:/Sitepages" or ConvertTo-ListGuid -FileLocation "C:/Sitepages/Home.aspx"

.NOTES

.LINK


Github repo: https://github.com/clymb3r/PowerShell
#>


  param(
    [Parameter(Mandatory = $true)]
    $FileLocation
  )
  
  if (Test-Path $FileLocation -PathType container) {

    $files = Get-ChildItem $FileLocation
    if ($files.count -gt 0)
    {
      foreach ($file in $files)
      {
        if ($file -ne $null) {
          try {

            $htmlContent = Get-Content $file.FullName;
            $htmlContent = Set-ListGuidContent ($htmlContent);
            $htmlContent > $file.FullName
          }
          catch {
            $message = "[-] Failed to convert file to ListId: $($file.FullName) "
            Write-Error $message
          }
        }

      }
    }

  }

  else
  {
    $htmlContent = Get-Content $FileLocation;
    if ($htmlContent -ne $null) {
      try {
        $htmlContent = Set-ListGuidContent ($htmlContent);
        $htmlContent > $FileLocation
      }
      catch {

        $message = "[-] Failed to convert file to ListId: $FileLocation "
        Write-Error $message

      }
    }

  }



}

function Set-ListGuidContent ($htmlContent)
{
  #Function that  uses the $ListMappings array to replace the contents of the file from a list name to a list guid.
  $htmlContent = $htmlContent.replace("ListName","ListID")

  if ($listsMappings.count -gt 0) {
    foreach ($Item in $listsMappings.GetEnumerator())
    {

      $htmlContent = $htmlContent.replace("ListDisplayName=`"$($Item.Key)`"","ListName=`"{$($Item.Value)}`"")
      $htmlContent = $htmlContent.replace("DefaultValue=`"$($Item.Key)`"","DefaultValue=`"$($Item.Value)`"")

    }
  }
  else
  {
    $message = "[-] Please configure the ListMappings variable before running the script "
    Write-Warning $message
    exit
  }
  return $htmlContent

}

