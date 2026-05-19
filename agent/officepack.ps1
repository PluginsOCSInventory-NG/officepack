<#
.Synopsis
----------------------------------------------------------
 Plugin for OCS Inventory NG 2.x
 Script :		Retrieve Microsoft Office informations
 Version :		2.25
 Date :		19/07/2023
 Author :		Creative Commons BY-NC-SA 3.0
 Author :		Nicolas DEROUET (nicolas.derouet[gmail]com)
 Contributor :		Stéphane PAUTREL (acb78.com), Eduardo Mozart de Oliveira, Antoine ROBIN
----------------------------------------------------------
 OS checked [X] on	32b	64b	(Professionnal edition)
	Windows XP	    [X]
	Windows Vista	[X]	[X]
	Windows 7	    [X]	[X]
	Windows 8.1	    [X]	[X]	
	Windows 10	    [X]	[X]
	Windows 2k8R2		[X]
	Windows 2k12R2		[X]
	Windows 2k16		[X]
 ---------------------------------------------------------
 Note :	No checked on Windows 8
 Included :	Office 2016 and 365 versions
 ---------------------------------------------------------

.Description
Retrieve Microsoft Office informations

#>

###
# Functions
###
Function EnumKey
{
    param
    (
      [UInt32]
      $hDefKey,

      [String]
      $sSubKeyName
    )

    $keys = (Invoke-CimMethod -ClassName StdRegProv -MethodName EnumKey -Arguments $PSBoundParameters).sNames
    Return $keys
}

Function GetBinaryValue
{
    param
    (
        [UInt32]
        $hDefKey,

        [String]
        $sSubKeyName,

        [String]
        $sValueName
    )

    $value = (Invoke-CimMethod -ClassName StdRegProv -MethodName GetBinaryValue -Arguments $PSBoundParameters).uValue

    if({GetBinaryValue -hDefKey $hDefKey -sSubKeyName $sSubKeyName -sValueName $sValueName} -ine $null)
    {
        Return $value
    }
    else
    {
        Return ''
    }
}

Function GetStringValue
{
    param
    (
        [UInt32]
        $hDefKey,

        [String]
        $sSubKeyName,

        [String]
        $sValueName
    )

    $value = (Invoke-CimMethod -ClassName StdRegProv -MethodName GetStringValue -Arguments $PSBoundParameters).sValue

    if({GetStringValue -hDefKey $hDefKey -sSubKeyName $sSubKeyName -sValueName $sValueName} -ine $null)
    {
        Return $value
    }
    else
    {
        Return ''
    }
}

Function getOfficeOSPPInfos
{
    Param (
        [string]
        $version
    )

    $objOfficeDict = @()
    $objOffice = @{}
    $index = 0
    $path = 'C:\Program Files (x86)\Microsoft Office\Office' + $version + '\OSPP.VBS'  
   
    if (Test-Path $path -PathType Leaf) {
        Write-Host "The path $path exist."
    }else {
        Write-Host "The path $path does not exist."
        $path = 'C:\Program Files\Microsoft Office\Office' + $version + '\OSPP.VBS'
    } 
    cscript.exe  $path /dstatus > $env:USERPROFILE\output.txt
    If (Test-Path $env:USERPROFILE\output.txt -PathType Leaf)
    {
        $data = Get-Content $env:USERPROFILE\output.txt
        Foreach ($str in $data)
        {
            $key = $str.Split(':')[0]
            $value = $str.Split(':')[1]
            switch ($key)
            {
                'PRODUCT ID'
                {
                    $productID = 1
                    $writeProduct = 1
                    $oProdID = $value.Trim()
                    $objOffice.Add('ProductID', $oProdID)
                }
                'SKU ID'
                {
                    $SKUID = 1
                    $writeProduct = 1
                    $oGUID = $value.Trim()
                    $objOffice.Add('SKUID', $oGUID)
                }
                'LICENSE NAME'
                {
                    $LicenseName = 1
                    $writeProduct = 1
                    $oProd = $value.Trim()
                    $objOffice.Add('LicenseName', $oProd)
                }
                'LICENSE DESCRIPTION'
                {
                    $LicenseDescription = 1
                    $writeProduct = 1
                    $oVer = $value.Trim()
                    $objOffice.Add('LicenseDescription', $oVer)
                }
                'ERROR DESCRIPTION'
                {
                    $writeProduct = 1
                    $oNote = $value.Trim()
                    $objOffice.Add('ErrorDescription', $oNote)
                }
                'Last 5 characters of installed product key'
                {
                    $ProductKey = 1
                    $writeProduct = 1
                    $oKey = 'XXXXX-XXXXX-XXXXX-XXXXX-' + $value.Trim()
                    $objOffice.Add('PartialProductKey', $oKey)
                }
            }
            If (
                ($SKUID -eq 1) -and
                ($LicenseName -eq 1) -and
                ($LicenseDescription -eq 1) -and
                ($ProductKey -eq 1)
            )
            {
                $productID = 0
                $SKUID = 0
                $LicenseName = 0
                $LicenseDescription = 0
                $ProductKey = 0
                $objOfficeDict += $objOffice
                $objOffice = @{}
            }
        }
    }
    Return $objOfficeDict
}

Function schKey
{
    param (
        [string]$regKey,

        [string]$likeOS,

        [int]$version
    )

    $aOSPPVersions = $aOffID[$version][1].Split('.')
    $oOfficeOSPPInfos = getOfficeOSPPInfos -version $aOSPPVersions[0]
    If ($oOfficeOSPPInfos.Length -ine 0)
    {
        If ($oOfficeOSPPInfos[0].Length -gt 0)
        {
            $oProdID = $oOfficeOSPPInfos[0]['ProductID']
            $oSKUID = $oOfficeOSPPInfos[0]['SKUID']
            If ($oOfficeOSPPInfos[0]['ErrorDescription'])
            {
                $oNote = $oOfficeOSPPInfos[0]['ErrorDescription']
            }
            Else
            {
                $oNote = ''
            }
            $oKey = $oOfficeOSPPInfos[0]['PartialProductKey']
        }
        Else
        {
            $oProdID = $oOfficeOSPPInfos['ProductID']
            $oSKUID = $oOfficeOSPPInfos['SKUID']
            If ($oOfficeOSPPInfos['ErrorDescription'])
            {
                $oNote = $oOfficeOSPPInfos['ErrorDescription']
            }
            Else
            {
                $oNote = ''
            }
            $oKey = $oOfficeOSPPInfos['PartialProductKey']
        }
    }
    If ($oProdID.Length -eq 0)
    {
        $oProdID = 'No ProductID'
    }
    $oVer = $aOffID[$version][0]
    If($oSKUID.Length -gt 0)
    {
        $oGUID = $oSKUID
    }
    $oInstall = '0'
    If ($osType -eq 32)
    {
        $wow = 'WOW6432Node'
    }
    Else
    {
        $wow = ''
    }
    If ($oProd.Length -eq 0)
    {
        $path = 'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall'
        $aUninstallKeys = EnumKey -hDefKey '0x80000002' -sSubKeyName $path
        If ($aUninstallKeys.Length -gt 0)
        {
            Foreach ($UninstallKey in $aUninstallKeys)
            {
                $path = 'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall\' + $UninstallKey
                $sValue = GetStringValue -hDefKey '0x80000002' -sSubKeyName $path -sValueName 'UninstallString'
                If ($sValue.Length -gt 0)
                {
                    If (
                        $sValue.ToLower().contains('microsoft office ' + $aOffID[$version][1].Substring(0,2)) -or
                        $sValue.ToLower().contains('productstoremove=') -gt 0
                    )
                    {
                        $oNote = $sValue.Substring($sValue.IndexOf('productstoremove='))
                        $oNote = $oNote -replace('productstoremove=','')
                        $oNote = $oNote.Substring(0, $oNote.IndexOf('_'))
                        If ($oNote.Contains('.'))
                        {
                            $oNote = $oNote.Substring(0, $oNote.IndexOf('.'))
                        }
                        $path = 'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall\' + $UninstallKey
                        $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName $path -sValueName 'DisplayName'
                        If ($oVer -eq '2016')
                        {
                            If ($oProd -notmatch 'Microsoft Office [\s\S]+? ' + $oVer)
                            {
                                $oVer = '2019'
                            }
                        }
                        $oInstall = '1'
                    }
                }
            }
        }
    }
    If ($oProd.Length -eq 0)
    {
        $Prod = GetStringValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'ProductName'
        If ($oProd.Length -gt 0)
        {
            $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'ConvertToEdition'
        }
    }
    If (
        $oProd.Length -eq 0 -and
        $oOfficeOSPPInfos.Length -gt 0
    )
    {
        If ($oOfficeOSPPInfos[0] -ne $null -and $oOfficeOSPPInfos.Length -gt 0) {
            $oProd = $oOfficeOSPPInfos[0]['LicenseName']
        }
        Else{
            $oProd = $oOfficeOSPPInfos['LicenseName']
        }
    }
    If ($oProd.Length -eq 0)
    {
        $oProd = 'Unidentifiable Office ' + $oVer
    }
    If ($oProdID -ine 'No ProductID')
    {
        If ($oProdID.SubString(7, 3) -eq 'OEM')
        {
            $oProd = $oProd + ' OEM'
        }
    }
    If ($oProd -match 'Office' -and $oProd -notmatch ('Unidentifiable Office ' + [regex]::Escape($oVer)))
    {
        writeXML -oVer $oVer -oProd $oProd -oProdID $oProdID -oBit $osType -oGUID $oGUID -oInstall $oInstall -oKey $oKey -oNote $oNote
    }
}

Function writeXML
{
    Param(
        [Parameter(Mandatory)]
        [string]$oVer,

        [string]$oProd,

        [string]$oProdID,

        [string]$oBit,

        [string]$oGUID,

        [string]$oInstall,

        [string]$oKey,

        [string]$oNote
    )
    $xml = ''
    $xml += "<OFFICEPACK>`n"
	$xml += "<OFFICEVERSION>" + $oVer + "</OFFICEVERSION>`n"
	$xml += "<PRODUCT>" + $oProd + "</PRODUCT>`n"
	$xml += "<PRODUCTID>" + $oProdID + "</PRODUCTID>`n"
	$xml += "<TYPE>" + $oBit + "</TYPE>`n"
	$xml += "<OFFICEKEY>" + $oKey + "</OFFICEKEY>`n"
	$xml += "<GUID>" + $oGUID + "</GUID>`n"
	$xml += "<INSTALL>" + $oInstall + "</INSTALL>`n"
	$xml += "<NOTE>" + $oNote + "</NOTE>`n"
	$xml += "</OFFICEPACK>`n"
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
    [Console]::WriteLine($xml)
}

###
# core
###
<#
Supported Office Families:
    Office 2000 -> KB230848 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/230848
    Office XP -> KB302663 - https://www.betaarchive.com/wiki/index.php?title=Microsoft_KB_Archive/302663
    Office 2003 -> KB832672 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/826217
    Office 2007 -> KB928516 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/928516
    Office 2010 -> KB2186281 - https://support.microsoft.com/en-us/topic/description-of-the-numbering-scheme-for-product-code-guids-in-office-2010-cceaef56-3d3f-1cae-8577-b4de3beaacfa
    Office 2013 -> KB2786054 - https://support.microsoft.com/help/2786054 
    Office 2016, O365 -> https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/numbering-scheme-for-product-guid
#>
$aOffID = @(
    @('XP', '10.0'),
    @('2003', '11.0'),
    @('2007', '12.0'),
    @('2010', '14.0'),
    @('2013', '15.0'),
    @('2016', '16.0')
)
$osType = 32
if ({GetStringValue -hDefKey '0x80000002' -sSubKeyName 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -sValueName 'PROCESSOR_ARCHITECTURE'} -eq 'AMD64')
{
    $osType = 64
}
$wow = ''
if ($osType = '64')
{
    $wow = 'WOW6432Node\'
}
For ($i = 0; $i -lt $aOffID.Length; $i++)
{
    $path = 'SOFTWARE + $wow + Microsoft\Office\' + $aOffID[$i][1] + '\Registration'
    schKey -regKey $path -likeOS False -version $i
}
